# Draft 1 version done
#
# Module Assessment Overview script
# Create 'Module Assessment Overview' Module Header block at position 0 (Published)
# then create Module Assessment Overview wikipage (Published)
# and insert this page into the newly created header block within every course
# Also generate error log to show any missed courses

import requests, glob, json, openpyxl, os
import pandas as pd
from tqdm import tqdm
from datetime import datetime
start = datetime.now()

with open(os.path.expanduser("~") + r'/testconfig.json') as json_data_file:
    configuration = json.load(json_data_file)
    access_token = configuration["canvas"]["access_token"]
    baseUrl = "https://"+configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization' : 'Bearer ' + access_token}
print(baseUrl)
csv_file = glob.glob('*.csv')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')

# Read HTML content from file
with open("page.html", "r", encoding="utf-8") as f:
    page_html = f.read()
#############################

def modulecreator(course_id):

    name = 'Module Assessment Overview'
    nameurl = f"{baseUrl}{course_id}/modules?module[name]={name}&module[position]=0"
    n = requests.post(nameurl, headers=header)

    if n.status_code == requests.codes.ok:
        data = n.json()
        print('Created: %s' % (data['name']))
        return data['id']
    else:
        print('Failed to create module for course %s' % course_id)
        return None


def create_module_assessment_page(course_id, html_file):

    with open(html_file, 'r', encoding='utf-8') as f:
        html = f.read()

    url = f"{baseUrl}{course_id}/pages"
    payload = {
        "wiki_page": {
            "title": "Module Assessment Overview",
            "body": html,
            "published": True
        }
    }
    
    n = requests.post(url, headers=header, json=payload)
    if n.status_code == requests.codes.ok:
        data = n.json()
        print('Created wiki page in course %s: %s' % (course_id, data['url']))
        return data['url']
    else:
        print('Failed to create wiki page in course %s - %s' % (course_id, n.status_code))
        return None


def insert_page_into_module(course_id, module_id, page_url):

    url = f"{baseUrl}{course_id}/modules/{module_id}/items"
    payload = {
        "module_item": {
            "type": "Page",
            "page_url": page_url,
            "published": True
        }
    }
    
    n = requests.post(url, headers=header, json=payload)
    if n.status_code == requests.codes.ok:
        data = n.json()
        print('Inserted page into module for course %s: module_item_id %s' % (course_id, data['id']))
    else:
        print('Failed to insert page into module for course %s - %s' % (course_id, n.status_code))

    return

def publish_module(course_id, module_id):

    url = f"{baseUrl}{course_id}/modules/{module_id}"
    payload = {"module": {"published": True}}
    n = requests.put(url, headers=header, json=payload)

    if n.status_code == requests.codes.ok:
        data = n.json()
        print('Published module %s in course %s' % (data['name'], course_id))
    else:
        print('Failed to publish module %s in course %s - %s' % (module_id, course_id, n.status_code))

    return


def main():

    html_file_paths = glob.glob('*.html')
    if len(html_file_paths) != 1:
        raise ValueError('should be only one HTML file in the current directory')
    else:
        html_file = html_file_paths[0]
    print("Filename:", html_file)

    error_log = []

    with tqdm(total=len(list(df.iterrows()))) as pbar:
        for index, row in df.iterrows():

            print("\n")
            course_id = row['course_id']

            module_id = modulecreator(course_id)
            if not module_id:
                base_host_url = "https://" + configuration["canvas"]["host"]
                course_url = f"{base_host_url}/courses/{course_id}"
                error_log.append({'course_id': course_id, 'URL': course_url})
                pbar.update(1)
                continue

            page_url = create_module_assessment_page(course_id, html_file)
            if not page_url:
                pbar.update(1)
                continue

            insert_page_into_module(course_id, module_id, page_url)
            publish_module(course_id, module_id)

            pbar.update(1)

    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"module_assessment_overview_error_log_{timestamp}.xlsx"
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
        print("\nError log saved as: %s" % ssname)
    else:
        print("\nNo errors encountered.")

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)