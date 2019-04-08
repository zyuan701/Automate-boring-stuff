# Automate-boring-stuff
Automate boring stuff in reading pdf and filling in forms
## Checking files under selected folder.vb
- This code will pop up a msgbox, allowing us to select the folder that we would like to check all files inside.
## Create Empty txt based on cells.vb
- This code will batch-create blank txt files under a selected folder, based on cells value in Excel.
## Illustration to FindFunction in VBA.vb
- Situtation: Visa Services Agents usually need to check if they have collected all requried supporting documents, as specified by
gov checklist.
- Solution: We create 2 worksheets under 1 workbook, ""FileName" and "Checklist". The documents might have slightly different names,
here, are assumed to contain basic keywords like "Passport", "COE", etc. As long as those keywords are detected, it will be written down
to Column B under sheet"FileName" whilst the undetected name will be shown as "N/A" and be formatted as Red font.
## CheckCategories_usingFind.vb
- The main feature of this codes enables the user to check if the keywords in Column"D" has ever appeared in another sheet's Column "A". Particluarly, in each cell of Column "D", you can enter mutilple keywords that you would like check and separate them by ";". (BTW, it's case insensitive)
- This solved the issue where people might name files differently but would like to categorize them according to a list of usual naming habit. For instance, "passport" or "ppt" in your file name will both be detected and categorized as "Passport Category" whilst "mum" or "dad" will be allocated to parents category.
## Batch Download.vb
- This is written to batch download files based on Urls in cells and then save it under Downloads folder. 
- Scenario: when collecting syllabus for students to apply for credit transfer, we could go to government website to download it based on course code. e.g.  The urls follow a same pattern as "https://training.gov.au/TrainingComponentFiles/BSB/BSBXXXXXX_R1.pdf", where BSBXXXXXX is the course code. 
