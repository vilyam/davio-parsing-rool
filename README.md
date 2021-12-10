Parsing tool to read the .doc/.docx files and extract the valuable data, transform it and load as a CSV.


Extract:
can read 'in' directory and iterate over all word doc files.


Transform:
The main value of the current script is declension of cases: from Genitive to Nominative for names and surnames.
Also, script can transliterate cyrillic words to Deutsch.
There are a few optimizaions such a replacement cache, and some predefined exclusions for grammatic rules.


Load:
There are files at the output: 
'out/out.csv' - contains correctly parsed and transformed values
'out/out_warnings.csv' - contains values that can't be parsed automaticlly and needs to be prepared manually.

----
The job was done under a special request from the The State Archives of Vinnytsia Region https://davio.gov.ua/ like a volunteer project.
