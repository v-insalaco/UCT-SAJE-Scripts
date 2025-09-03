import json
import requests
import os
import urllib.parse
import csv

# ------------------------------------------------------------
# Unpublish Canvas pages by title for a list of course IDs
# - Course IDs: CSV (1 column, no header)
# - Page names: TXT (one per line)
# ------------------------------------------------------------

# Load Canvas API configuration
config_path = os.path.expanduser("~") + r'/testconfig.json'
with open(config_path, encoding='utf-8') as json_data_file:
    configuration = json.load(json_data_file)

access_token = configuration["canvas"]["access_token"]
base_url = "https://" + configuration["canvas"]["host"].rstrip("/") + "/api/v1/"
headers = {'Authorization': 'Bearer ' + access_token}

def extract_course_id(raw_value: str) -> str:
    """
    Accepts a string from CSV, trims it, and validates it's numeric.
    Raises ValueError if invalid.
    """
    cid = (raw_value or "").strip()
    if cid.isdigit():
        return cid
    raise ValueError(f"Invalid course_id (must be digits only): {raw_value!r}")

def unpublish_page(course_id, page_url_slug):
    """
    Unpublish a Canvas page by its slug (page['url']).
    """
    encoded_slug = urllib.parse.quote(page_url_slug, safe='')
    url = f"{base_url}courses/{course_id}/pages/{encoded_slug}"
    payload = {"wiki_page": {"published": False}}
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"✅ Unpublished page: {page_url_slug}")
    else:
        print(f"❌ Failed to unpublish page {page_url_slug}: {response.status_code} - {response.text}")

def find_pages_by_names(course_id, page_names):
    """
    Finds pages whose titles match any of the provided names (case-insensitive) and unpublishes them.
    Handles pagination via Link headers.
    """
    url = f"{base_url}courses/{course_id}/pages"
    matching_pages = []
    # Normalize names for case-insensitive match
    target_names = {name.strip().lower() for name in page_names if name.strip()}

    params = {'per_page': 100}
    while url:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            print(f"❌ Failed to fetch pages for course {course_id}: {response.status_code} - {response.text}")
            return

        pages = response.json()
        for page in pages:
            title = (page.get('title') or '').strip().lower()
            if title in target_names:
                matching_pages.append(page)

        # Follow next page if provided in Link header
        next_link = response.links.get('next', {}).get('url')
        url = next_link
        params = None  # subsequent requests use the fully-qualified next URL

    if matching_pages:
        print(f"✅ Found {len(matching_pages)} matching page(s) in course {course_id}:")
        for page in matching_pages:
            print(f"- {page.get('title')} (URL: {page.get('html_url')})")
            if 'url' in page:
                unpublish_page(course_id, page['url'])
            else:
                print(f"⚠️ Skipping unpublish (no slug found) for {page}")
    else:
        print(f"⚠️ No matching pages found in course {course_id}.")

def load_lines_from_txt(file_path):
    """
    Reads non-empty trimmed lines from a .txt file (for page names).
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"⚠️ File not found: {file_path}")
        return []

def load_course_ids_from_csv(csv_path):
    """
    Reads the first column from a headerless CSV as course IDs.
    """
    course_ids = []
    try:
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue
                raw = row[0]
                try:
                    cid = extract_course_id(raw)
                    course_ids.append(cid)
                except ValueError as e:
                    print(f"⚠️ Skipping row with invalid course_id: {raw!r} ({e})")
    except FileNotFoundError:
        print(f"⚠️ File not found: {csv_path}")
    return course_ids

def main():
    print(">> Unpublish Pages (CSV course IDs) – v2025-08-22")

    # Prompts (CSV for course IDs, TXT for page names)
    csv_file_path = input("Enter path to .csv file with Canvas course IDs (1 column, no header):\n").strip()
    page_names_file_path = input("Enter path to .txt file with page names (one per line):\n").strip()

    course_ids = load_course_ids_from_csv(csv_file_path)
    page_names = load_lines_from_txt(page_names_file_path)

    if not course_ids:
        print("⚠️ No valid course IDs found in CSV.")
        return
    if not page_names:
        print("⚠️ No valid page names found in file.")
        return

    # Deduplicate while preserving order
    seen = set()
    unique_course_ids = []
    for cid in course_ids:
        if cid not in seen:
            seen.add(cid)
            unique_course_ids.append(cid)

    for course_id in unique_course_ids:
        find_pages_by_names(course_id, page_names)

if __name__ == "__main__":
    main()
