## Final/All Colleges/V1
#
# Post a 10 year future-dated Announcement for use in each course.
# Announcements are likely be different for each College.
# Using html file in same folder to create message.
# Also generate error log to show where this has been unsuccessful.
#

import requests, glob, json, os, openpyxl
import pandas as pd
from tqdm import tqdm
from datetime import datetime, timezone, timedelta
start = datetime.now()

with open(os.path.expanduser("~") + r'/testconfig.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseurl = "https://"+configuration["canvas"]["host"]+ "/api/v1/courses/"
header = {'Authorization' : 'Bearer ' + access_token}
payload = {}
csv_file = glob.glob('*.csv')
if len(csv_file) != 1:
    raise ValueError('should be only one file in the current directory')
csvfilename = csv_file[0]
pd.options.display.max_rows = None
df = pd.read_csv(csvfilename, encoding = 'unicode_escape')
print(df)
###########

def announcer(course_id, title, html_file_path):

    with open(html_file_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    # 1 Create announcement with a future post_at date
    url = baseurl + f"{course_id}/discussion_topics"

    # Time calcs - 10 years from now, at midnight UTC
    now = datetime.now(timezone.utc)
    future_date_obj = datetime(now.year + 10, now.month, now.day, 0, 0, 0, tzinfo=timezone.utc)
    future_date_obj = future_date_obj - timedelta(hours=1)
    future_date = future_date_obj.isoformat().replace('+00:00', 'Z')

    url = baseurl + f"{course_id}/discussion_topics"
    create_params = {
        "title": title,
        "message": html_content,
        "is_announcement": True,
        "delayed_post_at": future_date}
    
    r = requests.post(url, headers=header, params=create_params)
    if r.status_code == requests.codes.ok:
        data = r.json()
        print(f'Announcing {data["title"]} in {course_id}')

        # 2 Also change the course settings for announcements
        settingsurl = baseurl + f"{course_id}/settings"
        settings_payload = {
            "show_announcements_on_home_page": True,
            "home_page_announcement_limit": 1}
        
        sr = requests.put(settingsurl, headers=header, params=settings_payload)
        if sr.status_code == requests.codes.ok:
            print('Settings updated in %s' % (course_id))
        else:
            print('Settings update failed in %s' % (course_id))
        return True

    else:
        print('Cannot create announcement in %s: %s - %s' % (course_id, r.status_code, r.text))
        return False

def main():

    error_log = []

    html_file_paths = glob.glob('*.html')
    if len(html_file_paths) != 1:
        raise ValueError('should be only one HTML file in the current directory')
    else:
        html_file_path = html_file_paths[0]
    print("Filename:", html_file_path)

    title = 'Assignment Name'
    print("Title:", title)
    print("\n")

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for index, row in df.iterrows(): 
            print("\n")
            course_id = row['course_id']

            success = announcer(course_id, title, html_file_path)
            
            if not success:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
            pbar.update(1)

    # Error log
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"announcement_error_log_{timestamp}.xlsx"
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
            ws.column_dimensions[column_letter].width = max_length + 3

        wb.save(ssname)
        print("\nError log saved as: %s" % ssname)
    else:
        print("\nNo errors encountered.")

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)