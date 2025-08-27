# Final/All Colleges/V1
#
# Logic changed to adjust page_names list and Canvas pages to both use URL slugs.
# More accurate comparison, and further logic added to be agnostic of '-2' or other Canvas duplicates.
# Finds all 'base' matches, regardless of trailing number copies returned from Canvas API
#
# Find all wikipages and then find whatever Module Assessment information pages exist
# in all College/School courses by wikipage name.
# Identify naming conventions picked up from template discovery task,
# and Unpublish matching wikipages in a list of courses.
# Also generate error log to show where this has been unsuccessful.

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

def load_page_names_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return []

# Slug helpers
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

def find_pages_by_names(course_id, page_names, error_log):

    matching_pages = []
    url = f"{baseUrl}{course_id}/pages"
    params = {'per_page': 100}

    while url:
        response = requests.get(url, headers=header, params=params)

        if response.status_code != requests.codes.ok:
            print(f"Failed to fetch pages in {course_id}: {response.status_code} - {response.text}")
            base_host_url = "https://" + configuration["canvas"]["host"]
            course_url = f"{base_host_url}/courses/{course_id}"
            error_log.append({
                'course_id': course_id,
                'URL': course_url,
                'Error': f"{response.status_code} - {response.text}"
            })
            return None

        pages = response.json()

        ##### Find Logic #####
        for page in pages:
            # Getting Canvas Pages returned from API into same slug format as slugified page_names list
            canvas_slug = strip_numbers_from_canvas_pages(page['url'])
            if canvas_slug in page_names:
                matching_pages.append(page)

        # Handle pagination
        url = None
        if 'link' in response.headers:
            links = response.headers['link'].split(',')
            for link in links:
                if 'rel="next"' in link:
                    url = link[link.find('<')+1:link.find('>')]
                    break
    
    return matching_pages

def unpublish_page(course_id, page_url):
    url = f"{baseUrl}{course_id}/pages/{page_url}"
    payload = {"wiki_page": {"published": False}}

    r = requests.put(url, headers=header, json=payload)
    if r.status_code == requests.codes.ok:
        print(f"Unpublished: {page_url} in {course_id}")
        print('\n')
    else:
        print(f"Failed to unpublish page {page_url} in {course_id}: {r.status_code} - {r.text}")

def main():

    error_log = []

    input_file_paths = glob.glob('*.txt')
    if len(input_file_paths) != 1:
        raise ValueError('should be only one txt file in the current directory')
    else:
        file_path = input_file_paths[0]
    print("Filename:", file_path)

    # Build slugs upfront from the text file
    page_names = [slugify_page_list(name) for name in load_page_names_from_file(file_path)]

    if not page_names:
        print("No valid page names found in file.")
        return
    else:
        for i, name in enumerate(page_names, 1):
            print(f"{i}: {name}")

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for index, row in df.iterrows():
            course_id = row['course_id']
            print('\n')
            print(f"\nChecking course {course_id}...")

            try:
                matches = find_pages_by_names(course_id, page_names, error_log)

                if matches is None:
                    print(f"Course: {course_id} doesn't exist!")
                    pbar.update(1)
                    continue

                if matches:
                    print(f"Found {len(matches)} matching page(s) in {course_id}:")
                    print('\n')
                    for page in matches:
                        print(f"Page Found: {page['title']}")
                        unpublish_page(course_id, page['url'])
                else:
                    print(f"No matching pages found in {course_id}")

            except Exception as e:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({
                    'course_id': course_id,
                    'URL': course_url,
                    'Error': str(e)
                })
                print(f"Error in course {course_id}: {e}")

            pbar.update(1)

    # Save error log if needed
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"unpublish_module_assessment_error_log_{timestamp}.xlsx"
        df_errors.to_excel(ssname, header=True, index=False)

        # Auto column width formatting
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
        print(f"\nError log saved as: {ssname}")
    else:
        print("\nNo errors encountered for log.")

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)