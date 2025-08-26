import json
import requests
import os
import urllib.parse
import csv

# Load Canvas API configuration
config_path = os.path.expanduser("~") + r'/testconfig.json'
with open(config_path) as json_data_file:
    configuration = json.load(json_data_file)

access_token = configuration["canvas"]["access_token"]
base_url = "https://" + configuration["canvas"]["host"] + "/api/v1/"
headers = {'Authorization': 'Bearer ' + access_token}

# Error tracking
courses_with_no_matches = []
courses_with_failed_unpublishes = {}

def unpublish_page(course_id, page_url):
    encoded_page_url = urllib.parse.quote(page_url, safe='')
    url = f"{base_url}courses/{course_id}/pages/{encoded_page_url}"
    payload = {"wiki_page": {"published": False}}
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"✅ Unpublished page: {page_url}")
        return True
    else:
        print(f"❌ Failed to unpublish page {page_url}: {response.status_code} - {response.text}")
        return False

def find_pages_by_names(course_id, page_names):
    url = f"{base_url}courses/{course_id}/pages"
    matching_pages = []
    params = {'per_page': 100}
    while url:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            print(f"❌ Failed to fetch pages for course {course_id}: {response.status_code} - {response.text}")
            return
        pages = response.json()
        for page in pages:
            for name in page_names:
                if page['title'].strip().lower() == name.strip().lower():
                    matching_pages.append(page)
                    break
        url = None
        if 'link' in response.headers:
            links = response.headers['link'].split(',')
            for link in links:
                if 'rel="next"' in link:
                    url = link[link.find('<')+1:link.find('>')]
                    break
    if matching_pages:
        print(f"✅ Found {len(matching_pages)} matching page(s) in course {course_id}:")
        failed_pages = []
        for page in matching_pages:
            print(f"- {page['title']} (URL: {page['html_url']})")
            success = unpublish_page(course_id, page['url'])
            if not success:
                failed_pages.append(page['title'])
        if failed_pages:
            courses_with_failed_unpublishes[course_id] = failed_pages
    else:
        print(f"⚠️ No matching pages found in course {course_id}.")
        courses_with_no_matches.append(course_id)

def load_lines_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"⚠️ File not found: {file_path}")
        return []

def load_course_ids_from_csv(file_path):
    course_ids = []
    try:
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0].strip().isdigit():
                    course_ids.append(row[0].strip())
    except FileNotFoundError:
        print(f"⚠️ CSV file not found: {file_path}")
    return course_ids

def main():
    csv_file_path = input("Enter path to CSV file with Canvas course IDs:\n").strip()
    page_names_file_path = input("Enter path to .txt file with page names:\n").strip()

    course_ids = load_course_ids_from_csv(csv_file_path)
    page_names = load_lines_from_file(page_names_file_path)

    if not course_ids:
        print("⚠️ No valid course IDs found in CSV file.")
        return
    if not page_names:
        print("⚠️ No valid page names found in file.")
        return

    for course_id in course_ids:
        find_pages_by_names(course_id, page_names)

    # Error summary
    print("\n🔍 Error Summary:")
    if courses_with_no_matches:
        print("Courses with no matching pages:")
        for cid in courses_with_no_matches:
            print(f"- Course ID: {cid}")
    else:
        print("✅ All courses had matching pages.")

    if courses_with_failed_unpublishes:
        print("\nCourses with failed unpublishing attempts:")
        for cid, pages in courses_with_failed_unpublishes.items():
            print(f"- Course ID: {cid}")
            for page in pages:
                print(f"  • Page: {page}")
    else:
        print("✅ All matched pages were successfully unpublished.")

if __name__ == "__main__":
    main()
