# Final/All Colleges/V4
#
# Selective content import — imports ONLY the specified Assignment Group and its
# Assignments from the source course into every course listed in the CSV, then
# positions that Assignment Group at the top of the Assignments page.
# Writes an xlsx error log for any course where this did not succeed.
#
import glob
import json
import os
import time
from datetime import datetime

import openpyxl
import pandas as pd
import requests
from tqdm import tqdm

# ── Config ────────────────────────────────────────────────────────────────────

GROUP_NAME = "Assignment Templates"
PARENT_COURSE = 83108

POLL_INTERVAL = 10          # seconds between poll cycles
POLL_BACKOFF_MAX = 60       # cap on the backoff interval
MIGRATION_TIMEOUT = 45 * 60 # give up on a migration after this many seconds
REQUEST_RETRIES = 3         # transient-failure retries per request

TERMINAL_OK = {"completed"}
TERMINAL_BAD = {"failed"}

start = datetime.now()

with open(os.path.join(os.path.expanduser("~"), "config.json")) as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    base_host_url = "https://" + configuration["canvas"]["host"]
    baseurl = base_host_url + "/api/v1/"

header = {"Authorization": "Bearer " + access_token}
session = requests.Session()
session.headers.update(header)

csv_file = glob.glob("*.csv")
if len(csv_file) != 1:
    raise ValueError("should be only one csv file in the current directory")
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding="utf-8-sig")
print(df)


# ── HTTP helpers ──────────────────────────────────────────────────────────────

def request_with_retry(method, url, **kwargs):
    """
    Wraps a request with retries on connection errors, 5xx, and 403 rate limits.
    Returns the Response, or None if every attempt failed.
    """
    delay = 2
    for attempt in range(1, REQUEST_RETRIES + 1):
        try:
            r = session.request(method, url, timeout=60, **kwargs)
        except requests.RequestException as exc:
            if attempt == REQUEST_RETRIES:
                print(f"  Request error on {url}: {exc}")
                return None
            time.sleep(delay)
            delay *= 2
            continue

        # 403 from Canvas is usually throttling, not a permissions problem.
        if r.status_code >= 500 or r.status_code == 403:
            if attempt == REQUEST_RETRIES:
                return r
            time.sleep(delay)
            delay *= 2
            continue

        return r
    return None


def get_paginated(url, params=None):
    """Follows Canvas Link headers and returns the concatenated list, or None."""
    params = dict(params or {})
    params.setdefault("per_page", 100)
    results = []

    while url:
        r = request_with_retry("GET", url, params=params)
        if r is None or r.status_code != requests.codes.ok:
            status = "no response" if r is None else f"{r.status_code} - {r.text}"
            print(f"  Failed GET {url}: {status}")
            return None

        results.extend(r.json())

        url = r.links.get("next", {}).get("url")
        params = None  # the next link already carries the query string

    return results


# ── Source course lookup ──────────────────────────────────────────────────────

def course_name(parent_course):
    r = request_with_retry("GET", f"{baseurl}courses/{parent_course}")
    if r is None or r.status_code != requests.codes.ok:
        status = "no response" if r is None else f"{r.status_code} - {r.text}"
        print(f"Failed to fetch course_id: {status}")
        return None
    name = r.json()["name"]
    print(f"Course: {parent_course} is: {name}")
    return name


def get_source_group_and_assignments(parent_course, group_name):
    """
    Finds the target Assignment Group in the source course and the assignments
    inside it. Returns (group_id, [assignment_ids]) or (None, None).
    """
    groups = get_paginated(f"{baseurl}courses/{parent_course}/assignment_groups")
    if groups is None:
        return None, None

    group_id = None
    for group in groups:
        if group["name"].strip().lower() == group_name.strip().lower():
            group_id = group["id"]
            break

    if group_id is None:
        print(f"Assignment Group '{group_name}' not found in source course {parent_course}.")
        return None, None

    # Nested endpoint — the flat /assignments endpoint ignores assignment_group_id.
    assignments = get_paginated(
        f"{baseurl}courses/{parent_course}/assignment_groups/{group_id}/assignments"
    )
    if assignments is None:
        return group_id, None

    return group_id, [a["id"] for a in assignments]


# ── Migration ─────────────────────────────────────────────────────────────────

def create_content_migration(course_id, parent_course, source_group_id, source_assignment_ids):
    url = f"{baseurl}courses/{course_id}/content_migrations"

    # Canvas expects arrays of asset ids here, not {"id": "1"} hashes.
    payload = {
        "migration_type": "course_copy_importer",
        "settings": {"source_course_id": str(parent_course)},
        "select": {
            "assignment_groups": [str(source_group_id)],
            "assignments": [str(aid) for aid in source_assignment_ids],
        },
    }

    r = request_with_retry("POST", url, json=payload)
    if r is None:
        print(f"Failed to create migration in {course_id} - no response")
        return None, "no response from Canvas"

    if not r.ok:
        print(f"Failed to create migration in {course_id} - Error: {r.status_code} - {r.text}")
        return None, f"create failed: HTTP {r.status_code}"

    migration = r.json()
    print(f"Migration job created in course {course_id} with ID {migration['id']}.")
    return migration["id"], None


def check_migration_status(course_id, migration_id):
    """
    Returns (workflow_state, request_ok).
    request_ok is False when the request itself failed, so a network blip is not
    mistaken for a failed migration.
    """
    url = f"{baseurl}courses/{course_id}/content_migrations/{migration_id}"
    r = request_with_retry("GET", url)
    if r is None or r.status_code != requests.codes.ok:
        status = "no response" if r is None else f"{r.status_code} - {r.text}"
        print(f"  Could not check migration status for {course_id}: {status}")
        return None, False
    return r.json().get("workflow_state"), True


def move_assignment_group_to_top(course_id, group_name):
    groups = get_paginated(f"{baseurl}courses/{course_id}/assignment_groups")
    if groups is None:
        return False, "could not list assignment groups"

    for group in groups:
        if group["name"].strip().lower() == group_name.strip().lower():
            patch_url = f"{baseurl}courses/{course_id}/assignment_groups/{group['id']}"
            # Positions are 1-indexed. 0 is not the top, it is invalid.
            r = request_with_retry("PUT", patch_url, json={"position": 1})
            if r is not None and r.ok:
                print(f"Moved Assignment Group '{group_name}' to top spot in {course_id}.")
                return True, None
            status = "no response" if r is None else f"{r.status_code} - {r.text}"
            print(f"Failed to move Assignment Group in {course_id}: {status}")
            return False, "reposition failed"

    print(f"Assignment Group '{group_name}' not found in {course_id} after import.")
    return False, "group not present after import"


# ── Error log ─────────────────────────────────────────────────────────────────

def write_error_log(error_log):
    df_errors = pd.DataFrame(error_log)
    timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
    ssname = f"assignments_error_log_{timestamp}.xlsx"
    df_errors.to_excel(ssname, header=True, index=False)

    wb = openpyxl.load_workbook(ssname)
    ws = wb.active
    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = max_length + 5
    wb.save(ssname)
    print("\nError log saved as: %s" % ssname)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    course_name(PARENT_COURSE)

    print("\n--- Querying source course for selective import ---\n")
    source_group_id, source_assignment_ids = get_source_group_and_assignments(
        PARENT_COURSE, GROUP_NAME
    )

    if source_group_id is None or source_assignment_ids is None:
        print("Could not resolve source content. Aborting.")
        return

    if not source_assignment_ids:
        print(f"Assignment Group '{GROUP_NAME}' contains no assignments. Aborting.")
        return

    print(f"Source Assignment Group ID: {source_group_id}")
    print(f"Source Assignment IDs: {source_assignment_ids}")
    print(f"Total assignments to import: {len(source_assignment_ids)}\n")

    error_log = []

    def log_error(course_id, reason):
        error_log.append({
            "course_id": course_id,
            "reason": reason,
            "URL": f"{base_host_url}/courses/{course_id}",
        })

    # ── Phase 1: fire off every migration ──
    print("\n--- Starting all migrations ---\n")
    pending = {}   # course_id -> {"migration_id": ..., "started": ...}

    for _, row in df.iterrows():
        raw_id = row["course_id"]
        if pd.isna(raw_id):
            print("Skipping blank course_id row in CSV.")
            continue
        course_id = int(raw_id)  # guards against pandas float coercion

        migration_id, reason = create_content_migration(
            course_id, PARENT_COURSE, source_group_id, source_assignment_ids
        )
        if migration_id:
            pending[course_id] = {"migration_id": migration_id, "started": time.time()}
        else:
            log_error(course_id, reason)

    if not pending:
        print("\nNo migrations were created.")
        if error_log:
            write_error_log(error_log)
        return

    # ── Phase 2: poll until every migration resolves ──
    total = len(pending) + len(error_log)
    print(f"\n--- All {len(pending)} migrations launched. Polling for completion ---\n")

    pbar = tqdm(total=total, initial=len(error_log))
    interval = POLL_INTERVAL

    while pending:
        made_progress = False

        for course_id in list(pending.keys()):
            entry = pending[course_id]
            status, request_ok = check_migration_status(course_id, entry["migration_id"])

            if not request_ok:
                # Transient — leave it in pending and try again next cycle.
                continue

            elapsed = time.time() - entry["started"]

            if status in TERMINAL_OK:
                moved, reason = move_assignment_group_to_top(course_id, GROUP_NAME)
                if not moved:
                    log_error(course_id, reason)
                del pending[course_id]
                pbar.update(1)
                made_progress = True

            elif status in TERMINAL_BAD:
                print(f"Migration failed for {course_id}.")
                log_error(course_id, "migration reported failed")
                del pending[course_id]
                pbar.update(1)
                made_progress = True

            elif status == "waiting_for_select":
                # The select payload was not accepted; this will never finish.
                print(f"Migration for {course_id} is waiting_for_select — select payload rejected.")
                log_error(course_id, "stuck in waiting_for_select")
                del pending[course_id]
                pbar.update(1)
                made_progress = True

            elif elapsed > MIGRATION_TIMEOUT:
                print(f"Migration for {course_id} timed out in state '{status}'.")
                log_error(course_id, f"timed out in state '{status}'")
                del pending[course_id]
                pbar.update(1)
                made_progress = True

            # otherwise: pre_processing / queued / running -> keep waiting

        if pending:
            # Back off when nothing is resolving, to stay under the rate limit.
            interval = POLL_INTERVAL if made_progress else min(interval * 2, POLL_BACKOFF_MAX)
            time.sleep(interval)

    pbar.close()

    if error_log:
        write_error_log(error_log)
    else:
        print("\nNo errors encountered.")


if __name__ == "__main__":
    main()
    print("Running Time:", datetime.now() - start)
