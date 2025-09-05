# Final/All Colleges/V3
# Now working, and specifically finding just the assignments!
#
# Content import from course with only Assignment Group and 5 Assignment Templates in for speed.
# Imports into all courses in list.
# Then Positions this Assignment Group at the top of the Assignments section. 
# Also generates error log to show where this has been unsuccessful.

import requests, glob, json, os, time, openpyxl
import pandas as pd
from tqdm import tqdm
from datetime import datetime
start = datetime.now()
#############################

with open(os.path.expanduser("~") + r'/betaconfig.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseUrl = "https://"+configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization' : 'Bearer ' + access_token}
csv_file = glob.glob('*.csv')
if len(csv_file) != 1:
    raise ValueError('Should be only one CSV file in the current directory!')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')
print(df)

error_log = []
#############################

def xerox(parentcourse_id, course_id):
    """
    Import all assignments (and their assignment groups) from parent_course_id
    into course_id. Returns (success, migration_id).
    """

    # Create the content migration
    posturl = (
        f"{baseUrl}{course_id}/content_migrations"
        f"?migration_type=course_copy_importer"
        f"&settings[source_course_id]={parentcourse_id}"
        f"&selective_import=true"
    )
    preq = requests.post(posturl, headers=header)

    if preq.status_code == 404:
        print(f"[ERROR] Course {course_id} not found (404). Skipping...")
        return False, None
    if preq.status_code != requests.codes.ok:
        print(f"[ERROR] Failed to create migration for course {course_id}: {preq.status_code}")
        return False, None

    postdata = preq.json()
    migration_id = postdata.get('id')
    if not migration_id:
        print(f"[ERROR] No migration ID returned for course {course_id}")
        return False, None

    print(f"\nMigration created for course {course_id} with ID {migration_id}")

    # Select all assignments for import
    puturl = f"{baseUrl}{course_id}/content_migrations/{migration_id}?copy[all_assignments]=1"
    putreq = requests.put(puturl, headers=header)

    if putreq.status_code != requests.codes.ok:
        print(f"[ERROR] Failed to select assignments for course {course_id}: {putreq.status_code}")
        return False, migration_id

    putdata = putreq.json()
    workflow_state = putdata.get('workflow_state', 'unknown')
    print(f"Assignments-only import selected. Workflow state: {workflow_state}")

    return True, migration_id

def wait_for_migration(course_id, migration_id):

    print(f"Waiting for migration in {course_id}...")
    time.sleep(10)

    url = f"{baseUrl}{course_id}/content_migrations/{migration_id}"
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

def move_assignment_group_to_top(course_id, group_name):

    time.sleep(1)

    url = f"{baseUrl}{course_id}/assignment_groups"
    r = requests.get(url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"Failed to fetch assignment groups: {r.status_code} - {r.text}")
        return

    groups = r.json()
    for group in groups:
        if group["name"].strip().lower() == group_name.strip().lower():
            group_id = group["id"]
            patch_url = f"{baseUrl}{course_id}/assignment_groups/{group_id}"
            payload = {"position": 0}
            patch_r = requests.put(patch_url, headers=header, json=payload)
            if patch_r.status_code == requests.codes.ok:
                print(f"Moved Assignment Group '{group_name}' to top spot in {course_id}.")
            else:
                print(f"Failed to move Assignment Group: {patch_r.status_code} - {patch_r.text}")
            return

def main():

    parentcourse_id = input('Please enter the parent_course id: (e.g. 83815):\n')
    group_name = "Assignment Templates"

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for _, row in df.iterrows():
            course_id = row['course_id']
            success, migration_id = xerox(parentcourse_id, course_id)

            if not success:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
                print('\n')
                pbar.update(1)
                continue

            migration_success = wait_for_migration(course_id, migration_id)
            if not migration_success:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
                print('\n')
                pbar.update(1)
                continue

            move_assignment_group_to_top(course_id, group_name)
            pbar.update(1)

    # Save error log if any
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"assignments_error_log_{timestamp}.xlsx"
        df_errors.to_excel(ssname, header=True, index=False)

        # Auto-adjust column widths
        wb = openpyxl.load_workbook(ssname)
        ws = wb.active
        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = max_length + 5
        wb.save(ssname)
        print(f"\nError log saved as: {ssname}")
    else:
        print("\nNo errors encountered.")

if __name__ == "__main__": main()

end = datetime.now()
total_time = end - start
print("Running Time:", total_time)