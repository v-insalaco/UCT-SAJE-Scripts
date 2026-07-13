# Rename, Unpublish, Clean-Up, Announcement Deletion & Bulk Unpublish Script
#
# For every course listed in the CSV this script will:
#   1. Rename any "Module Assessment Overview" page -> "Module Assessment Overview 25_26"
#   2. Unpublish the renamed page
#   3. Unpublish any pages whose names match those listed in a .txt file
#      (slug-matched, agnostic of Canvas '-2' style duplicate suffixes)
#   4. Delete the module titled "Module Assessment Overview"
#   5. Delete the assignment group titled "Assignment Templates"
#      (and all assignments within it)
#   6. Delete all announcements with a delayed_post_at date of 18 Dec 2035 at 23:00
#
# Pages (Steps 1-3) are resolved in a SINGLE paginated fetch per course.
# Uses the same API access pattern, slug-matching logic, and pagination
# handling as the original scripts.

import requests, glob, json, openpyxl, os, re, unicodedata
import pandas as pd
from tqdm import tqdm
from datetime import datetime

start = datetime.now()

with open(os.path.expanduser("~") + r'/betaconfig.json') as json_data_file:
    configuration = json.load(json_data_file)

access_token = configuration["canvas"]["access_token"]
baseUrl = "https://" + configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization': 'Bearer ' + access_token}

print(baseUrl)

csv_file = glob.glob('*.csv')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')

# ---------- Configuration ----------

OLD_PAGE_NAME = "Module Assessment Overview"
NEW_PAGE_NAME = "Module Assessment Overview 25_26"
TARGET_MODULE_NAME = "Module Assessment Overview"
TARGET_ASSIGNMENT_GROUP_NAME = "Assignment Templates"
TARGET_ANNOUNCEMENT_DATE = "2035-12-18T23:00"   # 18 Dec 2035 at 23:00

# ---------- Slug helpers ----------

def slugify_page_list(name: str) -> str:
    """Convert a page name to a Canvas-style slug."""
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    name = name.strip().lower()
    name = re.sub(r"[^a-z0-9]+", "-", name)
    name = name.strip("-")
    return name

def strip_numbers_from_canvas_pages(slug: str) -> str:
    """Strip numeric suffix from a Canvas slug if present."""
    return re.sub(r"-\d+$", "", slug)

# Pre-compute the target slug once
target_slug = slugify_page_list(OLD_PAGE_NAME)

# ---------- Helper: build course URL for error logging ----------

def course_url_for(course_id):
    base_host_url = "https://" + configuration["canvas"]["host"]
    return f"{base_host_url}/courses/{course_id}"

# ---------- Helper: check if a datetime string matches the target date ----------

def matches_target_date(date_string):
    """Check if a Canvas ISO 8601 datetime string matches 18 Dec 2035 at 23:00."""
    if not date_string:
        return False
    # Canvas returns dates like "2035-12-18T23:00:00Z" or with timezone offset.
    # We compare against the first 16 characters: "2035-12-18T23:00"
    return date_string[:16] == TARGET_ANNOUNCEMENT_DATE

# ---------- Generic paginated GET ----------

def paginated_get(url, params=None, error_context=None, error_log=None):
    """Yield items from a paginated Canvas API endpoint."""
    if params is None:
        params = {}
    params.setdefault('per_page', 100)

    while url:
        response = requests.get(url, headers=header, params=params)
        if response.status_code != requests.codes.ok:
            if error_log is not None and error_context is not None:
                error_log.append({**error_context,
                                  'Error': f"{response.status_code} - {response.text}"})
            return
        for item in response.json():
            yield item

        # Handle pagination
        url = None
        params = {}          # params only needed on first request
        if 'link' in response.headers:
            links = response.headers['link'].split(',')
            for link in links:
                if 'rel="next"' in link:
                    url = link[link.find('<')+1:link.find('>')]
                    break

# ==========================================================
# .TXT FILE LOADER
# ==========================================================

def load_page_names_from_file(file_path):
    """Load page names from a .txt file, one per line."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return []

# ==========================================================
# SINGLE PAGE FETCH — resolves Steps 1, 2 & 3 in one pass
# ==========================================================

def fetch_and_partition_pages(course_id, page_names_slugs, error_log):
    """Fetch all pages for a course in one paginated pass and return two
    separate lists:

        rename_targets  — pages whose base slug matches target_slug
                          (used for rename + unpublish, Steps 1 & 2)
        txt_targets     — pages whose base slug appears in page_names_slugs
                          (used for bulk unpublish, Step 3)

    A page can appear in both lists if its slug satisfies both conditions,
    though in practice the .txt list and the rename target are distinct.
    """
    rename_targets = []
    txt_targets = []
    err_ctx = {'course_id': course_id, 'URL': course_url_for(course_id)}

    for page in paginated_get(f"{baseUrl}{course_id}/pages",
                              error_context=err_ctx, error_log=error_log):
        base_slug = strip_numbers_from_canvas_pages(page['url'])

        if base_slug == target_slug:
            rename_targets.append(page)

        if base_slug in page_names_slugs:
            txt_targets.append(page)

    return rename_targets, txt_targets

# ==========================================================
# RENAME a page
# ==========================================================

def rename_page(course_id, page, error_log):
    """Rename a page to NEW_PAGE_NAME via PUT."""
    page_url_slug = page['url']
    url = f"{baseUrl}{course_id}/pages/{page_url_slug}"
    payload = {'wiki_page': {'title': NEW_PAGE_NAME}}

    response = requests.put(url, headers=header, json=payload)
    if response.status_code == 200:
        return True
    else:
        error_log.append({
            'course_id': course_id,
            'URL': course_url_for(course_id),
            'page_title': page['title'],
            'page_slug': page_url_slug,
            'Error': f"Rename failed: {response.status_code} - {response.text}"
        })
        return False

# ==========================================================
# UNPUBLISH a page
# ==========================================================

def unpublish_page(course_id, page_url_slug, error_log):
    """Set published=false on the given page."""
    url = f"{baseUrl}{course_id}/pages/{page_url_slug}"
    payload = {'wiki_page': {'published': False}}

    response = requests.put(url, headers=header, json=payload)
    if response.status_code == 200:
        return True
    else:
        error_log.append({
            'course_id': course_id,
            'URL': course_url_for(course_id),
            'page_slug': page_url_slug,
            'Error': f"Unpublish failed: {response.status_code} - {response.text}"
        })
        return False

# ==========================================================
# DELETE the "Module Assessment Overview" MODULE
# ==========================================================

def delete_target_module(course_id, error_log):
    """Find and delete all modules whose name matches TARGET_MODULE_NAME."""
    deleted = 0
    err_ctx = {'course_id': course_id, 'URL': course_url_for(course_id)}

    for module in paginated_get(f"{baseUrl}{course_id}/modules",
                                error_context=err_ctx, error_log=error_log):
        if module.get('name', '').strip() == TARGET_MODULE_NAME:
            del_url = f"{baseUrl}{course_id}/modules/{module['id']}"
            response = requests.delete(del_url, headers=header)
            if response.status_code == 200:
                deleted += 1
                print(f"    Deleted module \"{module['name']}\" (id {module['id']})")
            else:
                error_log.append({
                    'course_id': course_id,
                    'URL': course_url_for(course_id),
                    'module_name': module['name'],
                    'module_id': module['id'],
                    'Error': f"Module delete failed: {response.status_code} - {response.text}"
                })
    return deleted

# ==========================================================
# DELETE the "Assignment Templates" ASSIGNMENT GROUP
# ==========================================================

def delete_target_assignment_group(course_id, error_log):
    """Find and delete all assignment groups matching TARGET_ASSIGNMENT_GROUP_NAME.

    If move_assignments_to is omitted on the DELETE request, Canvas also
    deletes all assignments inside the group — which is the desired behaviour.
    """
    deleted = 0
    err_ctx = {'course_id': course_id, 'URL': course_url_for(course_id)}

    for group in paginated_get(f"{baseUrl}{course_id}/assignment_groups",
                               error_context=err_ctx, error_log=error_log):
        if group.get('name', '').strip() == TARGET_ASSIGNMENT_GROUP_NAME:
            del_url = f"{baseUrl}{course_id}/assignment_groups/{group['id']}"
            response = requests.delete(del_url, headers=header)
            if response.status_code == 200:
                deleted += 1
                print(f"    Deleted assignment group \"{group['name']}\" "
                      f"(id {group['id']}) and all its assignments")
            else:
                error_log.append({
                    'course_id': course_id,
                    'URL': course_url_for(course_id),
                    'group_name': group['name'],
                    'group_id': group['id'],
                    'Error': f"Assignment group delete failed: "
                             f"{response.status_code} - {response.text}"
                })
    return deleted

# ==========================================================
# DELETE announcements with delayed_post_at of 18 Dec 2035 23:00
# ==========================================================

def delete_target_announcements(course_id, error_log):
    """Find and delete all announcements whose delayed_post_at matches
    18 Dec 2035 at 23:00.

    Canvas announcements are discussion topics with is_announcement=true.
    We fetch them via the discussion_topics endpoint filtered to
    only_announcements, then check the delayed_post_at field.
    """
    deleted = 0
    err_ctx = {'course_id': course_id, 'URL': course_url_for(course_id)}

    for topic in paginated_get(
            f"{baseUrl}{course_id}/discussion_topics",
            params={'only_announcements': True, 'per_page': 100},
            error_context=err_ctx, error_log=error_log):

        delayed_date = topic.get('delayed_post_at', '')
        if matches_target_date(delayed_date):
            del_url = f"{baseUrl}{course_id}/discussion_topics/{topic['id']}"
            response = requests.delete(del_url, headers=header)
            if response.status_code == 200:
                deleted += 1
                print(f"    Deleted announcement \"{topic['title']}\" "
                      f"(id {topic['id']}, delayed_post_at: {delayed_date})")
            else:
                error_log.append({
                    'course_id': course_id,
                    'URL': course_url_for(course_id),
                    'announcement_title': topic.get('title', ''),
                    'announcement_id': topic['id'],
                    'Error': f"Announcement delete failed: "
                             f"{response.status_code} - {response.text}"
                })
    return deleted

# ==========================================================
# PRE-FLIGHT: Load & validate .txt page name list
# ==========================================================

input_file_paths = glob.glob('*.txt')
if len(input_file_paths) != 1:
    raise ValueError("There should be exactly one .txt file in the current directory.")

txt_file_path = input_file_paths[0]
print(f"\nLoading page names from: {txt_file_path}")

raw_page_names = load_page_names_from_file(txt_file_path)
if not raw_page_names:
    raise ValueError(f"No valid page names found in {txt_file_path}. Aborting.")

# Build the set of slugs once — used for O(1) lookup per page
page_names_slugs = {slugify_page_list(name) for name in raw_page_names}

print(f"Loaded {len(page_names_slugs)} unique slug(s) from .txt file:")
for i, slug in enumerate(sorted(page_names_slugs), 1):
    print(f"  {i}: {slug}")

# ==========================================================
# MAIN LOOP
# ==========================================================

error_log = []

total_renamed = 0
total_unpublished = 0
total_txt_unpublished = 0
total_modules_deleted = 0
total_groups_deleted = 0
total_announcements_deleted = 0
courses_processed = 0

for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing courses"):
    course_id = row.iloc[0]
    print(f"\n[{idx+1}/{len(df)}] Course {course_id}")

    # -- Single page fetch: partitions results for Steps 1, 2 & 3 --
    rename_targets, txt_targets = fetch_and_partition_pages(
        course_id, page_names_slugs, error_log
    )

    # -- Steps 1 & 2: Rename + Unpublish "Module Assessment Overview" pages --
    if rename_targets:
        for page in rename_targets:
            print(f"  [Steps 1 & 2] Found page: \"{page['title']}\" (slug: {page['url']})")
            renamed = rename_page(course_id, page, error_log)
            if renamed:
                total_renamed += 1
                print(f"    Renamed -> \"{NEW_PAGE_NAME}\"")
                # After rename, Canvas updates the slug to match the new title
                new_slug = slugify_page_list(NEW_PAGE_NAME)
                if unpublish_page(course_id, new_slug, error_log):
                    total_unpublished += 1
                    print(f"    Unpublished \"{NEW_PAGE_NAME}\"")
    else:
        print(f"  [Steps 1 & 2] No pages found matching \"{OLD_PAGE_NAME}\".")

    # -- Step 3: Unpublish pages from .txt list --
    if txt_targets:
        print(f"  [Step 3] Found {len(txt_targets)} page(s) matching .txt list:")
        for page in txt_targets:
            print(f"    Page found: \"{page['title']}\" (slug: {page['url']})")
            if unpublish_page(course_id, page['url'], error_log):
                total_txt_unpublished += 1
                print(f"    Unpublished: \"{page['title']}\"")
    else:
        print(f"  [Step 3] No pages matching .txt list found.")

    # -- Step 4: Delete module "Module Assessment Overview" --
    total_modules_deleted += delete_target_module(course_id, error_log)

    # -- Step 5: Delete assignment group "Assignment Templates" --
    total_groups_deleted += delete_target_assignment_group(course_id, error_log)

    # -- Step 6: Delete announcements dated 18 Dec 2035 23:00 --
    total_announcements_deleted += delete_target_announcements(course_id, error_log)

    courses_processed += 1

# ==========================================================
# SUMMARY & ERROR LOG
# ==========================================================

end = datetime.now()
elapsed = end - start

print("\n" + "=" * 60)
print("SUMMARY")
print("=" * 60)
print(f"  Courses processed          : {courses_processed}")
print(f"  Pages renamed              : {total_renamed}")
print(f"  Pages unpublished          : {total_unpublished}")
print(f"  Pages unpublished (.txt)   : {total_txt_unpublished}")
print(f"  Modules deleted            : {total_modules_deleted}")
print(f"  Assignment groups deleted  : {total_groups_deleted}")
print(f"  Announcements deleted      : {total_announcements_deleted}")
print(f"  Errors                     : {len(error_log)}")
print(f"  Time elapsed               : {elapsed}")
print("=" * 60)

if error_log:
    error_df = pd.DataFrame(error_log)
    error_filename = f"error_log_{end.strftime('%Y%m%d_%H%M%S')}.csv"
    error_df.to_csv(error_filename, index=False)
    print(f"\nError log saved to: {error_filename}")
else:
    print("\nNo errors encountered.")
