# Automate-boring-stuff
Automate boring stuff in reading pdf and filling in forms
## Illustration to FindFunction in VBA.vb
*Situtation: Visa Services Agents usually need to check if they have collected all requried supporting documents, as specified by
gov checklist.
*Solution: We create 2 worksheets under 1 workbook, ""FileName" and "Checklist". The documents might have slightly different names,
here, are assumed to contain basic keywords like "Passport", "COE", etc. As long as those keywords are detected, it will be written down
to Column B under sheet"FileName" whilst the undetected name will be shown as "N/A" and be formatted as Red font.