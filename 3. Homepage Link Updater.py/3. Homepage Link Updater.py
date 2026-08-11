# Final/All Colleges/V2
# 3. Homepage Link Updater
# 
# Find the homepage or front page of a course that's used as the chosen homepage,
# and on this identified page find the link/button link text from the template discovery task loaded via text file
# that corresponds with how this College, School or Department deals with Module Assessment/Assessment & Feedback.
# Change this to point to the Module Assessment Overview wikipage.
# Now checks whether module-assessment-overview-2 exists first; if so, links there instead.
# Error log created for invalid courses.

import requests, glob, json, os, re, openpyxl
import pandas as pd
from tqdm import tqdm
from datetime import datetime
from bs4 import BeautifulSoup
start = datetime.now()

with open(os.path.expanduser("~") + r'/config.json') as f:
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

#####################

def load_link_replacements(filename="link_replacements.txt"):
    replacements = {}
    with open(filename, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue  # allow comments and blank lines
            try:
                old, new = line.split(",", 1)
                replacements[old.strip()] = new.strip()
            except ValueError:
                print(f"⚠️ Skipping malformed line in {filename}: {line}")
    return replacements

LINK_REPLACEMENTS = load_link_replacements()

# Cache for page existence checks so we don't re-query the same course
_page_exists_cache = {}

def check_page_exists(course_id, page_slug):
    """Check whether a specific wiki page exists in a course. Results are cached."""
    cache_key = (course_id, page_slug)
    if cache_key in _page_exists_cache:
        return _page_exists_cache[cache_key]

    url = baseUrl + f"{course_id}/pages/{page_slug}"
    r = requests.get(url, headers=header)
    exists = r.status_code == requests.codes.ok
    _page_exists_cache[cache_key] = exists
    return exists

def resolve_assessment_page(course_id):
    """
    Determine which Module Assessment Overview page to link to for a given course.
    Returns the page slug to use: 'module-assessment-overview-2' if it exists,
    otherwise 'module-assessment-overview'.
    """
    preferred = "module-assessment-overview-2"
    fallback = "module-assessment-overview"

    if check_page_exists(course_id, preferred):
        print(f"   ✅ Page '{preferred}' found in course {course_id} — using it.")
        return preferred
    else:
        print(f"   ℹ️ Page '{preferred}' not found in course {course_id} — falling back to '{fallback}'.")
        return fallback

def extract_styles_from_element(element):
    font_size = None
    color = None
    is_bold = False

    def get_color_and_size(tag):
        fs = None
        col = None
        if tag.has_attr('style'):
            style_str = tag['style'].replace(" ", "").lower()
            size_match = re.search(r'font-size:([^;]+)', style_str)
            color_match = re.search(r'color:([^;]+)', style_str)
            if size_match:
                fs = size_match.group(1).strip()
            if color_match:
                col = color_match.group(1).strip()
        return fs, col

    # 1️⃣ Check direct children first (like <span> inside <a>)
    for child in element.find_all(recursive=False):
        fs, col = get_color_and_size(child)
        if font_size is None and fs:
            font_size = fs
        if color is None and col:
            color = col
        if font_size and color:
            break

    # 2️⃣ Check the element itself
    fs, col = get_color_and_size(element)
    if font_size is None and fs:
        font_size = fs
    if color is None and col:
        color = col

    # 3️⃣ Walk up parents if needed
    current = element.parent
    while current and (font_size is None or color is None):
        fs, col = get_color_and_size(current)
        if font_size is None and fs:
            font_size = fs
        if color is None and col:
            color = col
        current = current.parent

    # Bold check
    if element.find('strong') or element.find('b'):
        is_bold = True

    return font_size, color, is_bold

def update_links_in_html(html, course_id):
    soup = BeautifulSoup(html, 'html.parser')
    changed = False

    print(f"\n--- Checking links in course {course_id} HTML ---")

    # Resolve which assessment page to use for this course (checked once per course)
    target_page_slug = resolve_assessment_page(course_id)

    for a in soup.find_all("a", href=True):
        print(f"Checking link: {a['href']}")
        href_lower = a['href'].lower()
        for old_pattern, new_path in LINK_REPLACEMENTS.items():
            if old_pattern.lower() in href_lower:
                print(f"Matched old pattern: {old_pattern}")

                # Replace the page slug in the new_path with the resolved target
                # e.g. "/pages/module-assessment-overview" → "/pages/module-assessment-overview-2"
                adjusted_new_path = new_path.replace(
                    "module-assessment-overview",
                    target_page_slug
                )

                if adjusted_new_path.startswith("/"):
                    new_href = f"{webBaseUrl}{course_id}{adjusted_new_path}"
                elif adjusted_new_path.startswith("http"):
                    new_href = adjusted_new_path
                else:
                    new_href = f"{webBaseUrl}{course_id}/{adjusted_new_path}"

                parent_div = a.find_parent("div")
                font_size, color, is_bold = extract_styles_from_element(a)

                a['href'] = new_href

                style_parts = []
                if font_size:
                    style_parts.append(f"font-size: {font_size}")
                if color:
                    style_parts.append(f"color: {color}")
                if style_parts:
                    a['style'] = "; ".join(style_parts)

                # Set the display text based on which page we're linking to
                display_text = "Module Assessment Overview"

                if is_bold:
                    a.clear()
                    bold_tag = soup.new_tag("strong")
                    bold_tag.string = display_text
                    a.append(bold_tag)
                else:
                    a.string = display_text

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
    """Fetch course details and return default_view, or None if invalid."""
    course_url = baseUrl + f"{course_id}"
    r = requests.get(course_url, headers=header)
    if r.status_code != requests.codes.ok:
        print(f"⚠️ Failed to fetch course {course_id} details: {r.status_code}")
        return None  # ❌ don't append here — let main() handle logging

    course_data = r.json()
    return course_data.get('default_view', 'unknown')

def main():

    with tqdm(total=len(df), desc="Updating courses") as pbar:
        for _, row in df.iterrows():
            course_id = row['course_id']
            course_url = f"{webBaseUrl}{course_id}"
            print('\n')

            try:
                default_view = check_homepage_type(course_id)

                if default_view is None:
                    # Invalid course → log once and skip
                    error_log.append({'course_id': course_id, 'URL': course_url, 'Error': "Invalid course ID"})
                    pbar.update(1)
                    continue

                if default_view == 'wiki':
                    print(f"\nCalling update_front_page for course {course_id}")
                    success, msg, front_html_url = update_front_page(course_id)
                    if front_html_url:
                        course_url = front_html_url

                elif default_view == 'syllabus':
                    print(f"\nCalling update_syllabus for course {course_id}")
                    success, msg = update_syllabus(course_id)

                else:
                    success = False
                    msg = f"Unknown default_view: {default_view}"

                if not success:
                    error_log.append({'course_id': course_id, 'URL': course_url, 'Error': msg})

            except Exception as e:
                error_log.append({'course_id': course_id, 'URL': course_url, 'Error': str(e)})
                print(f"⚠️ Exception in course {course_id}: {e}")

            pbar.update(1)

    # Save error log
    if error_log:
        df_errors = pd.DataFrame(error_log)
        timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        ssname = f"homepage_link_updater_error_log_{timestamp}.xlsx"
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
