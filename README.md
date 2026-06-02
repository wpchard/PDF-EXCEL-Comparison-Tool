this file takes two specific reports for my work and identitifies the changes from day to day. 
This program can take either pdf files or excel files, converting the pdfs into excel files for use.
The set up for this current version is quite rigid, requiring a specifc amount of columns and order of columns.
The program tracks if there was a document in the old xls but not in the new(removed), if document is in new but not in old (new entry), or if there was a change in last scan date. (new scan). defaults to (unchanged)

Steps for use
1. choose the older report
2. choose the newer report
3. if either report was a pdf, the program will create a new excel file of the pdf document
4. wait for processing to complete
5. choose the name and where to save the newly created report

the output will include all changes in the newer report and display a column including the type of change 
that row had: unchanged, new scan, new entry, or removed.
a small summary will be generated below the table including totals of each type and total rows in the document.

NOTES:
This program removes headers and other dead rows in excel files if they are blank or do not start with a number; my rows started with a date.
Many functions revolve around there being 8 columns, and will add or remove if the value is different. change expected columns to your desired columns as well as any relevant functions.
The pdf conversion method I used manually sets the size of each column. If your pdf table is clearly seperated you can use that to break the columns. column size and names of the columns will change depending on the pdf.
If you are planning to use manual column sizes, the names of the columns do not particularly matter as a later step removed all header columns. however, there must be something there or it will break and keep everything in a single column. the first column CANNOT start with a number as the logic that removes headers and dead columns removes rows that do not start with numbers. 


