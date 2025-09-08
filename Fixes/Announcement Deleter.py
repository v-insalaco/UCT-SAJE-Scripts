# Announcements Deleter - V2
# Working? In Test currently

import requests, glob, json, os, openpyxl, time
import pandas as pd
from tqdm import tqdm
from datetime import datetime
start = datetime.now()

with open(os.path.expanduser("~") + r'/testconfig.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseurl = "https://" + configuration["canvas"]["host"] + "/api/v1/"
header = {'Authorization': 'Bearer ' + access_token}
csv_file = glob.glob('*.csv')
if len(csv_file) != 1:
    raise ValueError('should be only one file in the current directory')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')
print(df)

error_log = []
#############################

def delete_announcement_from_course(course_id, announcement_id):
    url = f"{baseurl}courses/{course_id}/discussion_topics/{announcement_id}"
    print("Delete URL: " + url)

    r = requests.delete(url, headers=header)

    if r.status_code == requests.codes.ok:
        print(f"Successfully deleted announcement {announcement_id} in course {course_id}")
    else:
        print(f"[ERROR] Failed to delete announcement {announcement_id} in {course_id}: {r.status_code}")
    return r

def list_announcements(course_id):
    announcements_found_thus_far = []
    url = f"{baseurl}courses/{course_id}/discussion_topics?only_announcements=1"
    print("List URL: " + url)

    r = requests.get(url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"[ERROR] Could not fetch announcements for course {course_id}: {r.status_code}")
        return []

    page_response = r.json()
    announcements_found_thus_far.extend(page_response)

    # Handle pagination
    while "next" in r.links:
        r = requests.get(r.links["next"]["url"], headers=header)
        page_response = r.json()
        announcements_found_thus_far.extend(page_response)

    return announcements_found_thus_far

def main():
    assignment_name = "Assignment Name"
    print('\n')

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for _, row in df.iterrows():
            course_id = row['course_id']
            announcements = list_announcements(course_id)

            found = False
            for announcement in announcements:
                if announcement["title"].strip().lower() == assignment_name.strip().lower():
                    announcement_id = announcement['id']
                    print(f"Deleting announcement '{announcement['title']}' (id={announcement_id}) from course {course_id}")
                    resp = delete_announcement_from_course(course_id, announcement_id)
                    if resp.status_code == requests.codes.ok:
                        found = True
                    else:
                        base_host_url = "https://" + configuration["canvas"]["host"]
                        course_url = f"{base_host_url}/courses/{course_id}"
                        error_log.append({
                            'course_id': course_id,
                            'announcement_id': announcement_id,
                            'title': announcement["title"],
                            'URL': course_url
                        })
            if not found:
                print(f"No matching announcement '{assignment_name}' found in course {course_id}")

            pbar.update(1)

    # Save error log if any
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"announcements_error_log_{timestamp}.xlsx"
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

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)