# 3. Homepage Link Updater
# Test
# Things to sort - don't need double error log!
# Seems to work alright in draft but sort error log
# 
# Find the homepage or front page of a course that’s used as the chosen homepage,
# and on this identified page find the link/button link text from the template discovery task
# that corresponds with how this College, School or Department deals with Module Assessment/Assessment & Feedback.
# Where this does not point towards our previously created Module Assessment Overview page (script 1),
# change this to point to the Module Assessment Overview wikipage.
# Error log too

import requests, glob, json, os
import pandas as pd
from tqdm import tqdm
from datetime import datetime
from bs4 import BeautifulSoup
import re
start = datetime.now()

with open(os.path.expanduser("~") + r'/testconfig.json') as f:
    configuration = json.load(f)
    access_token = configuration["canvas"]["access_token"]
    baseUrl = "https://" + configuration["canvas"]["host"] + "/api/v1/courses/"
header = {'Authorization': 'Bearer ' + access_token}
webBaseUrl = f"https://{configuration['canvas']['host']}/courses/"

print(baseUrl)
csv_file = glob.glob('*.csv')
csvfilename = csv_file[0]
df = pd.read_csv(csvfilename, encoding='unicode_escape')
error_log = []

LINK_REPLACEMENTS = {
    "/assignments": "/pages/module-assessment-overview",
    "/pages/assessment": "/pages/module-assessment-overview",
    "/pages/assignments-for-this-module": "/pages/module-assessment-overview",}

####################

def extract_styles_from_div(div_tag):
    font_size = None
    color = None
    is_bold = False

    if div_tag and div_tag.has_attr('style'):
        style_str = div_tag['style']
        size_match = re.search(r'font-size\s*:\s*([^;]+)', style_str, re.IGNORECASE)
        color_match = re.search(r'color\s*:\s*([^;]+)', style_str, re.IGNORECASE)
        if size_match:
            font_size = size_match.group(1).strip()
        if color_match:
            color = color_match.group(1).strip()

    if div_tag and (div_tag.find('strong') or div_tag.find('b')):
        is_bold = True

    return font_size, color, is_bold

def update_links_in_html(html, course_id):
    soup = BeautifulSoup(html, 'html.parser')
    changed = False

    print(f"\n--- Checking links in course {course_id} HTML ---")
    print(html)

    for a in soup.find_all("a", href=True):
        print(f"Checking link: {a['href']}")
        href_lower = a['href'].lower()
        for old_pattern, new_path in LINK_REPLACEMENTS.items():
            if old_pattern.lower() in href_lower:
                print(f"Matched old pattern: {old_pattern}")
                if new_path.startswith("/"):
                    new_href = f"{webBaseUrl}{course_id}{new_path}"
                elif new_path.startswith("http"):
                    new_href = new_path
                else:
                    new_href = f"{webBaseUrl}{course_id}/{new_path}"

                parent_div = a.find_parent("div")
                font_size, color, is_bold = extract_styles_from_div(parent_div)

                a['href'] = new_href

                style_parts = []
                if font_size:
                    style_parts.append(f"font-size: {font_size}")
                if color:
                    style_parts.append(f"color: {color}")
                if style_parts:
                    a['style'] = "; ".join(style_parts)

                if is_bold:
                    a.clear()
                    bold_tag = soup.new_tag("strong")
                    bold_tag.string = "Module Assessment Overview"
                    a.append(bold_tag)
                else:
                    a.string = "Module Assessment Overview"

                print(f"Updated link to: {a['href']} with style: {a.get('style', '')}")
                changed = True
                break
    return str(soup), changed

def update_front_page(course_id):
    """Fetch the front page HTML directly and update links if needed."""
    front_page_url = baseUrl + f"{course_id}/front_page"
    r = requests.get(front_page_url, headers=header)
    
    if r.status_code != requests.codes.ok:
        print(f"   URL: {front_page_url}")
        print(f"   Status: {r.status_code}")
        try:
            print(f"   Response: {r.json()}")
        except Exception:
            print(f"   Response (text): {r.text[:500]}")
        return False, f"Failed to fetch front page for course {course_id}", None
    
    page_data = r.json()
    front_html_url = page_data.get('html_url', None)
    html = page_data.get('body', '')

    print(f"Front Page Homepage found for course {course_id}")
    
    updated_html, changed = update_links_in_html(html, course_id)
    
    if changed:
        payload = {'wiki_page[body]': updated_html}
        p = requests.put(front_page_url,
                         headers={**header, 'Content-Type': 'application/x-www-form-urlencoded'},
                         data=payload)
        if p.status_code == requests.codes.ok:
            print(f"✅ Updated front page for course {course_id}")
            return True, None, front_html_url
        else:
            return False, f"Failed to update front page for course {course_id}: {p.status_code}", front_html_url
    
    return True, None, front_html_url  # no links changed

def update_syllabus(course_id):
    syllabus_url = baseUrl + f"{course_id}?include[]=syllabus_body"
    r = requests.get(syllabus_url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"   URL: {syllabus_url}")
        print(f"   Status: {r.status_code}")
        try:
            print(f"   Response: {r.json()}")
        except Exception:
            print(f"   Response (text): {r.text[:500]}")
        return False, f"Failed to fetch syllabus for course {course_id}"

    default_view = r.json().get('default_view')
    html = r.json().get('syllabus_body', '') or ''

    if default_view == 'syllabus':
        print(f"Syllabus Body Homepage found for course {course_id}")

    updated_html, changed = update_links_in_html(html, course_id)

    if changed:
        payload = {'course[syllabus_body]': updated_html}
        p = requests.put(baseUrl + f"{course_id}",
                         headers={**header, 'Content-Type': 'application/x-www-form-urlencoded'},
                         data=payload)
        if p.status_code == requests.codes.ok:
            print(f"✅ Updated syllabus body for course {course_id}")
            return True, None
        else:
            return False, f"Failed to update syllabus for course {course_id}: {p.status_code}"
    else:
        return True, None
    
def check_homepage_type(course_id):
    """Fetch course details and print default_view."""
    course_url = baseUrl + f"{course_id}"
    r = requests.get(course_url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"⚠️ Failed to fetch course {course_id} details: {r.status_code}")
        # Correct dictionary syntax for append
        error_log.append({'course_id': course_id, 'URL': f"{webBaseUrl}{course_id}"})
        return None

    course_data = r.json()
    default_view = course_data.get('default_view', 'unknown')
    print(f"Course {course_id} - default_view: {default_view}")

    return default_view

def main():

    with tqdm(total=len(df), desc="Updating courses") as pbar:
        for _, row in df.iterrows():
            course_id = row['course_id']
            course_url = f"{webBaseUrl}{course_id}"  # default fallback URL
            print('\n')

            try:
                default_view = check_homepage_type(course_id)

                if default_view == 'wiki':
                    print('\n')
                    print(f"Calling update_front_page for course {course_id}")
                    success, msg, front_html_url = update_front_page(course_id)
                    # Use front_html_url if returned
                    if front_html_url:
                        course_url = front_html_url

                elif default_view == 'syllabus':
                    print('\n')
                    print(f"Calling update_syllabus for course {course_id}")
                    success, msg = update_syllabus(course_id)

                else:
                    success = False
                    msg = f"Unknown default_view: {default_view}"

                if not success:
                    error_log.append({'course_id': course_id, 'URL': course_url})
                    if msg:
                        print(f"⚠️ {msg}")

            except Exception as e:
                error_log.append({'course_id': course_id, 'URL': course_url})
                print(f"⚠️ Exception in course {course_id}: {e}")

            pbar.update(1)

    # Save error log
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"update_links_error_log_{timestamp}.xlsx"
        df_errors.to_excel(ssname, header=True, index=False)
        print(f"\nError log saved as: {ssname}")
    else:
        print("\nNo errors encountered.")

if __name__ == "__main__": main()

total_time = datetime.now() - start
print("Running Time:", total_time)