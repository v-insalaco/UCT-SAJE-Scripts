# Attempt to fix the homepage issue occurred!
# Works to find only the default_view = modules entries, and change them to wiki (front page)

import os, json, glob, requests, openpyxl
import pandas as pd
from tqdm import tqdm
from datetime import datetime
start = datetime.now()

with open(os.path.expanduser("~") + r'/config.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseUrl = "https://" + configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization': 'Bearer ' + access_token}
print(baseUrl)

csv_file = glob.glob('*.csv')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')
error_log = []

#############################

def get_course_default_view(course_id):
    """Check the current default_view setting of the course."""
    url = baseUrl + str(course_id)
    r = requests.get(url, headers=header)
    if r.status_code == 200:
        return r.json().get('default_view')
    else:
        print(f"[ERROR] Could not fetch course {course_id}: {r.status_code}")
        error_log.append(f"[ERROR] {course_id}: {r.status_code}")
        return None

def set_course_home_to_frontpage(course_id):
    """Set the default_view of the course to 'wiki' (front page)."""
    url = baseUrl + str(course_id)
    payload = {"course": {"default_view": "wiki"}}
    r = requests.put(url, headers=header, json=payload)
    if r.status_code == 200:
        print(f"[SUCCESS] Homepage set to Front Page for course {course_id}")
    else:
        print(f"[ERROR] Failed to set homepage for course {course_id}: {r.status_code}")

#############################

def main():

    with tqdm(total=len(df), desc="Updating courses") as pbar:
        for _, row in df.iterrows():
            print("\n")
            course_id = row['course_id']
            base_host_url = "https://" + configuration["canvas"]["host"]
            course_url = f"{base_host_url}/courses/{course_id}"

            try:
                current_view = get_course_default_view(course_id)

                if current_view and current_view == "modules":
                    print(f"Course {course_id} default_view = {current_view}, updating...")
                    set_course_home_to_frontpage(course_id)

                elif current_view == "wiki":
                    print(f"Course {course_id} already has Front Page as homepage, skipping.")

            except Exception as e:
                error_log.append({'course_id': course_id, 'URL': course_url, 'Error': str(e)})
                print(f"Exception in course {course_id}: {e}")

            pbar.update(1)

    # Save error log if needed
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"update_homepage_error_log_{timestamp}.xlsx"
        df_errors.to_excel(ssname, header=True, index=False)

        # Excel file formatting
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
            adjusted_width = max_length + 5
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(ssname)
        print(f"\nError log saved as: {ssname}")
    else:
        print("\nNo errors encountered.")

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)