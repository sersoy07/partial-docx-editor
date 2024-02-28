# PARTIAL-DOCX-EDITOR
**PARTIAL-DOCX-EDITOR** is a python script to be used to add or replace the specific field in docx file.


## PURPOSE
**PARTIAL-DOCX-EDITOR** is a script written to create specific .docx and .pdf files from a .docx template. It is especially useful for Cover Letter and Resumes to create recruiter and/or company specific pdf files.


## BENEFITS
- No need to enter to file, change it and save as pdf.
- Generic output name can be determined and then company name and position will be added to output with date and time.


## INSTALLATION & HOW TO USE
- Download the code your system name and install the requirements or download *partial-docx-editor.exe*.
- Move/copy your Cover Letter and Resume Template docx file(ex:cv template.docx) to folder where *partial-docx-editor* is.
- Edit your template docx with \<Company_name\>, \<Job_title\>, \<Recruiter_name\> and \<Additional_field\>(you can remove/rename or remove these fields.)
- Run *partial-docx-editor*
- In the first run, the code will create *replacement_code.txt* file.
- Add your replacement codes as a new line.
- Continue to fill the fields in the cli.
```
Ex:
~~~~~~~~~ WELCOME TO PARTIAL DOCX EDITOR ~~~~~~~~~
!!! >>> File type of CV template must be in docx

Default : cv template
Enter the CV Word Template name without file type:

Default : Cover Letter and Resume
Output Filename: Cover Letter & Resume (Data A. Networkson)

!!! Don't forget to edit and match "replacement_code.txt" with "cv template.docx"

<Company_name>: Innovate Networks Inc.

<Job_title>:  Network Automation Engineer

Default : Recruiter
<Recruiter_name>: Allison M. Talentia

<Additional_field>: If you want to schedule a call, I am available after 14:00 on weekdays.
100%|██████████████████████████████████████████████████████████████████████████████████████████████████████████████████████████████| 1/1 [00:05<00:00,  5.99s/it]
'q' for quit.
```
- *Cover Letter & Resume (Data A. Networkson) for  Network Automation Engineer at Innovate Networks Inc._\<date\>_\<time\>.docx* and *Cover Letter & Resume (Data A. Networkson) for  Network Automation Engineer at Innovate Networks Inc._\<date\>_\<time\>.docx* in folder.
- You can continue to create new files or you can quit by *q*.