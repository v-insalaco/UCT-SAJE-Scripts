# Final/All Colleges/V3
# Note: Make sure to use a course with no other assignments in it!!
#
# Selective content import — only imports the specified Assignment Group and
# its Assignments from the source course. Imports into all courses in list.
# Then positions this Assignment Group at the top of the Assignments section.
# Also generates error log to show where this has been unsuccessful.
#
# V2: Fires off ALL migrations first, then polls them concurrently — no serial waiting.
# V3: Selective import — only brings in the target Assignment Group and its 5 Assignments.

import requests, glob, json, os, openpyxl, time
import pandas as pd
from tqdm import tqdm
from datetime import datetime

start = datetime.now()

with open(os.path.expanduser("~") + r'/config.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseurl = "https://" + configuration["canvas"]["host"] + "/api/v1/"
header = {'Authorization': 'Bearer ' + access_token}

csv_file = glob.glob('*.csv')
if len(csv_file) != 1:
    raise ValueError('should be only one file in the current directory')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='utf-8-sig')
print(df)

##########################

def course_name(parent_course):
    url = f"{baseurl}courses/{parent_course}"
    r = requests.get(url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"Failed to fetch course_id: {r.status_code} - {r.text}")
    else:
        parent_name = r.json()
        print(f"Course: {parent_course} is: {parent_name['name']}")


def get_source_group_and_assignments(parent_course, group_name):
    """
    Queries the source course to find the target Assignment Group and its
    Assignments. Returns (group_id, [assignment_ids]) or (None, None) on failure.
    """
    # Get assignment groups
    url = f"{baseurl}courses/{parent_course}/assignment_groups"
    r = requests.get(url, headers=header, params={"per_page": 100})
    if r.status_code != requests.codes.ok:
        print(f"Failed to fetch assignment groups from source: {r.status_code} - {r.text}")
        return None, None

    group_id = None
    for group in r.json():
        if group["name"].strip().lower() == group_name.strip().lower():
            group_id = group["id"]
            break

    if group_id is None:
        print(f"Assignment Group '{group_name}' not found in source course {parent_course}.")
        return None, None

    # Get assignments in that group
    url = f"{baseurl}courses/{parent_course}/assignments"
    params = {"assignment_group_id": group_id, "per_page": 100}
    r = requests.get(url, headers=header, params=params)
    if r.status_code != requests.codes.ok:
        print(f"Failed to fetch assignments from source: {r.status_code} - {r.text}")
        return group_id, None

    assignment_ids = [a["id"] for a in r.json()]
    return group_id, assignment_ids


def create_content_migration(course_id, parent_course, source_group_id, source_assignment_ids):
    url = f"{baseurl}courses/{course_id}/content_migrations"

    # Build selective import payload
    select = {
        "assignment_groups": {str(source_group_id): "1"},
        "assignments": {str(aid): "1" for aid in source_assignment_ids}
    }

    payload = {
        "migration_type": "course_copy_importer",
        "settings": {"source_course_id": parent_course},
        "select": select
    }

    r = requests.post(url, headers=header, json=payload)
    if r.status_code == requests.codes.ok:
        migration = r.json()
        print(f"Migration job created in course {course_id} with ID {migration['id']}.")
        return migration['id']
    else:
        print(f"Failed to create migration in {course_id} - Error: {r.status_code} - {r.text}")
        return None


def check_migration_status(course_id, migration_id):
    """Returns the workflow_state string, or None on request failure."""
    url = f"{baseurl}courses/{course_id}/content_migrations/{migration_id}"
    r = requests.get(url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"Failed to check migration status for {course_id}: {r.status_code} - {r.text}")
        return None
    return r.json()["workflow_state"]


def move_assignment_group_to_top(course_id, group_name):
    url = f"{baseurl}courses/{course_id}/assignment_groups"
    r = requests.get(url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"Failed to fetch assignment groups: {r.status_code} - {r.text}")
        return

    groups = r.json()
    for group in groups:
        if group["name"].strip().lower() == group_name.strip().lower():
            group_id = group["id"]
            patch_url = f"{baseurl}courses/{course_id}/assignment_groups/{group_id}"
            payload = {"position": 0}
            patch_r = requests.put(patch_url, headers=header, json=payload)
            if patch_r.status_code == requests.codes.ok:
                print(f"Moved Assignment Group '{group_name}' to top spot in {course_id}.")
            else:
                print(f"Failed to move Assignment Group: {patch_r.status_code} - {patch_r.text}")
            return


def main():
    group_name = "Assignment Templates"
    parent_course = 83108
    course_name(parent_course)

    # ── Pre-flight: Identify the source group and its assignments ──
    print("\n--- Querying source course for selective import ---\n")
    source_group_id, source_assignment_ids = get_source_group_and_assignments(parent_course, group_name)

    if source_group_id is None or source_assignment_ids is None:
        print("Could not resolve source content. Aborting.")
        return

    print(f"Source Assignment Group ID: {source_group_id}")
    print(f"Source Assignment IDs: {source_assignment_ids}")
    print(f"Total assignments to import: {len(source_assignment_ids)}\n")

    error_log = []
    base_host_url = "https://" + configuration["canvas"]["host"]

    # ── Phase 1: Fire off ALL migrations ──
    print("\n--- Starting all migrations ---\n")
    pending = {}  # {course_id: migration_id}

    for _, row in df.iterrows():
        course_id = row['course_id']
        migration_id = create_content_migration(course_id, parent_course, source_group_id, source_assignment_ids)
        if migration_id:
            pending[course_id] = migration_id
        else:
            error_log.append({
                'course_id': course_id,
                'URL': f"{base_host_url}/courses/{course_id}"
            })

    # ── Phase 2: Poll all migrations until every one resolves ──
    print(f"\n--- All {len(pending)} migrations launched. Polling for completion ---\n")

    pbar = tqdm(total=len(pending) + len(error_log), initial=len(error_log))

    while pending:
        for course_id in list(pending.keys()):
            migration_id = pending[course_id]
            status = check_migration_status(course_id, migration_id)

            if status == "completed":
                print(f"Migration complete for {course_id}.")
                move_assignment_group_to_top(course_id, group_name)
                del pending[course_id]
                pbar.update(1)

            elif status == "failed" or status is None:
                print(f"Migration failed for {course_id}.")
                error_log.append({
                    'course_id': course_id,
                    'URL': f"{base_host_url}/courses/{course_id}"
                })
                del pending[course_id]
                pbar.update(1)

            # Any other status (pre_processing, running, etc.) → leave in pending

        if pending:
            time.sleep(5)  # brief pause before next poll cycle

    pbar.close()

    # ── Error log ──
    if error_log:
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
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[column_letter].width = max_length + 5
        wb.save(ssname)
        print("\nError log saved as: %s" % ssname)
    else:
        print("\nNo errors encountered.")


if __name__ == "__main__":
    main()

total_time = datetime.now() - start
print("Running Time:", total_time)
