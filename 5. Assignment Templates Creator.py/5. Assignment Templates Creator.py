# Final/All Colleges/V1
# Note: Make sure to use a course with no other assignments in it!!
# 
#
# Content import from course with only Assignment Group and 5 Assignment Templates in for speed.
# Imports into all courses in list.
# Then Positions this Assignment Group at the top of the Assignments section. 
# Also generates error log to show where this has been unsuccessful.

import requests, glob, json, os, openpyxl, time
import pandas as pd
from tqdm import tqdm
from datetime import datetime
start = datetime.now()

with open(os.path.expanduser("~") + r'/betaconfig.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseurl = "https://"+configuration["canvas"]["host"]+ "/api/v1/"
header = {'Authorization' : 'Bearer ' + access_token}
csv_file = glob.glob('*.csv')
if len(csv_file) != 1:
    raise ValueError('should be only one file in the current directory')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding = 'unicode_escape')
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

    return

def create_content_migration(course_id, parent_course):

    url = f"{baseurl}courses/{course_id}/content_migrations"
    payload = {
        "migration_type": "course_copy_importer",
        "settings": {"source_course_id": parent_course}}

    r = requests.post(url, headers=header, json=payload)
    if r.status_code == requests.codes.ok:
        migration = r.json()
        print(f"Migration job created in course {course_id} with ID {migration['id']}.")
        return migration['id']
    else:
        print(f"Failed to create migration in {course_id} - Error: {r.status_code} - {r.text}")
        return None

def wait_for_migration(course_id, migration_id):

    print(f"Waiting for migration in {course_id}...")
    time.sleep(40) # no point checking straight away. But should we wait 40 seconds?

    url = f"{baseurl}courses/{course_id}/content_migrations/{migration_id}"
    while True:
        r = requests.get(url, headers=header)
        if r.status_code != requests.codes.ok:
            print(f"Failed to check migration status: {r.status_code} - {r.text}")
            return False

        migration = r.json()
        status = migration["workflow_state"]
        print(f"Migration {migration_id} status: {status}")

        if status == "completed":
            print("Migration complete.")
            return True
        elif status == "failed":
            print("Migration failed.")
            return False

        time.sleep(2)

def move_assignment_group_to_top(course_id, group_name):

    time.sleep(2)

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

    error_log = []

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for index, row in df.iterrows(): 
            print("\n")
            course_id = row['course_id']

            migration_id = create_content_migration(course_id, parent_course)
            if not migration_id:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
                continue

            migration_success = wait_for_migration(course_id, migration_id)
            if not migration_success:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
                continue

            move_assignment_group_to_top(course_id, group_name) 
            pbar.update(1)

    # Error log
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

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)