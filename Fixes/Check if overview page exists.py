import requests, glob, json, os, openpyxl
import pandas as pd
from tqdm import tqdm
from datetime import datetime
import re

start = datetime.now()

# ----------------------------
# Load Canvas config
# ----------------------------
with open(os.path.expanduser("~") + r'/config.json') as f:
    configuration = json.load(f)
    access_token = configuration["canvas"]["access_token"]
    baseUrl = "https://" + configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization': 'Bearer ' + access_token}
webBaseUrl = f"https://{configuration['canvas']['host']}/courses/"

print("Canvas API Base URL:", baseUrl)

# ----------------------------
# Load course list from CSV
# ----------------------------
csv_file = glob.glob('*.csv')
if not csv_file:
    raise FileNotFoundError("No CSV file found in current directory.")
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')

# ----------------------------
# Globals
# ----------------------------
error_log = []
results = []

TARGET_PAGE_URL = "module-assessment-overview-coss"

def page_exists(course_id):
    """
    Return True if a page exists with URL exactly TARGET_PAGE_URL.
    Ignores other variations like -2, -3, etc.
    """
    url = baseUrl + f"{course_id}/pages?per_page=100"
    session = requests.Session()
    session.headers.update(header)

    while url:
        r = session.get(url)
        if r.status_code != requests.codes.ok:
            return None, f"Failed to fetch pages ({r.status_code})"

        pages = r.json()
        for page in pages:
            page_url = page.get("url", "").strip()
            # Exact match
            if page_url == TARGET_PAGE_URL:
                return True, page.get("html_url", None)

        # Pagination
        url = r.links['next']['url'] if 'next' in r.links else None

    return False, None


# ----------------------------
# Main function
# ----------------------------
def main():
    with tqdm(total=len(df), desc="Checking courses") as pbar:
        for _, row in df.iterrows():
            course_id = row['course_id']
            course_url = f"{webBaseUrl}{course_id}"
            try:
                exists, page_url = page_exists(course_id)
                if exists is True:
                    results.append({
                        'course_id': course_id,
                        'URL': course_url,
                        'PageFound': True,
                        'PageURL': page_url
                    })
                elif exists is False:
                    results.append({
                        'course_id': course_id,
                        'URL': course_url,
                        'PageFound': False,
                        'PageURL': None
                    })
                else:
                    error_log.append({
                        'course_id': course_id,
                        'URL': course_url,
                        'Error': page_url  # page_url contains error message here
                    })
            except Exception as e:
                error_log.append({
                    'course_id': course_id,
                    'URL': course_url,
                    'Error': str(e)
                })
            pbar.update(1)

    # ----------------------------
    # Save results
    # ----------------------------
    timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")

    if results:
        df_results = pd.DataFrame(results)
        ssname = f"page_check_results_{timestamp}.xlsx"
        df_results.to_excel(ssname, header=True, index=False)
        print(f"\n✅ Results saved as: {ssname}")

    if error_log:
        df_errors = pd.DataFrame(error_log)
        ssname = f"page_check_errors_{timestamp}.xlsx"
        df_errors.to_excel(ssname, header=True, index=False)
        print(f"⚠️ Errors logged as: {ssname}")

# ----------------------------
# Run
# ----------------------------
if __name__ == "__main__":
    main()
    total_time = datetime.now() - start
    print("Running Time:", total_time)
