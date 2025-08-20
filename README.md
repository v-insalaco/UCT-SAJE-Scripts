Script Tasks
<br>1. Module Assessment Overview.py
<br>Create ’Module Assessment Overview’ Module Header block at position 0,
<br>then create Module Assessment Overview wikipage (Published),
<br>and insert this page into the newly created header block, and Publish the header block, within every course.
<br>Also generate error log to show where this has been unsuccessful.<br>
<br>a. Requirements: Course list for all Taught Modules, Module Assessment Overview page in identified course.
<br>b. Single Script: Not College/School specific.
<br>c. Likely Hit Rate: 100% of provided list/lists

<br>2. Unpublish Module Assessment Pages.py [Discovery]
<br>Find all wikipages in each targeted course,
<br>and then find whatever Module Assessment information pages exist in all College, or School courses by wikipage name.
<br>Identify naming conventions picked up from template discovery task,
<br>and Unpublish matching wikipages in a list of courses.
<br>Also generate error log to show where this has been unsuccessful.<br>
<br>a. Requirements: Course list, DE-led (potential ESO support) checks to identify College templates as well as local level templates where different in Schools or Departments for finding of templated page names.
<br>b. Complex Script: Potential for one script to build in various bits of logic to identify naming conventions for this page in many or all areas, although may prove tricky to implement if some contexts use more vague page naming conventions.
<br>c. Likely Hit Rate: Less than 100% of courses in real world, where page names are likely to vary. Error log can be created by script.

<br>3. Homepage Link Updater.py [Discovery]
<br>Find the homepage or front page of a course that’s used as the chosen homepage,
<br>and on this identified page find the link/button link text from the template discovery task
<br>that corresponds with how this College, School or Department deals with Module Assessment/Assessment & Feedback.
<br>Where this does not point towards our previously created Module Assessment Overview page (script 1),
<br>change this to point to the Module Assessment Overview wikipage.<br>
<br>a. Requirements: Course list, DE-led (potential ESO support) checks to identify College templates as well as local level templates where different in Schools or Departments for finding link text for how a College, School or Department deals with Module Assessment/Assessment & Feedback.
<br>b. Complex Script: Potential for one script to build in various bits of logic to identify link text in many or all areas, although may prove tricky to implement if some contexts use more generic link text with risks of changing non-desirable links. Most likely worth keeping area specific scripts to avoid risks/much higher stakes script for student experience.
<br>c. Likely Hit Rate: Less than 100% of courses in real world, where template usage is likely to vary for course homepages. Error log can be created by script.

<br>4. Announcement Poster.py [Different Text]
<br>Create an unpublished, undated Announcement for use in each course.
<br>Announcements are likely be different for each College, School or Department.
<br>Also generate error log to show where this has been unsuccessful.<br>
<br>a. Requirements: Course list, Announcement Title and Announcement text to be required for each setting an Announcement is to be created in.
<br>b. Multiple Scripts: College, School or Department specific.
<br>c. Likely Hit Rate: 100% of provided list/lists

<br>5. Assignment Templates Creator.py
<br>Create an ’Assignment Templates’ Assignment Group (unpublished),
<br>then into this Assignment Group, import the 5 Assignment Template wikipages (unpublished) within every course.
<br>Position this Assignment Group at the top of the Assignments section.
<br>Also generate error log to show where this has been unsuccessful.<br>
<br>a. Requirements: Course list for all Taught Modules, Assignment Template wikipages within course to copy from.
<br>b. Single Script: Not College/School specific.
<br>c. Likely Hit Rate: 100% of provided list/lists
