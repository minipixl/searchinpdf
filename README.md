# textsearchinPDF
OS:       Windows 11  
Use:      opens console to 1. enter path and 2. textstring to search for.  
Purpose:  search in all PDF files within specified path and all subdirectories.  
Output:   Console shows file name, page numbers of the matches and total number of matches.  
Output:   Excel file 'matches.xlsx' with the number of matches and hyperlinks to respective files.  

Enter path in console  
![Screenshot](./screenshots/4.jpg)  

Enter search string in console  
![Screenshot](./screenshots/5.jpg)  

See results in console with:  
no. of files found to search in  
path, file, page no. for each match  
total match count  
![Screenshot](./screenshots/6.jpg)  

Press 'enter' to start Excel or Crtl+c to abort  
![Screenshot](./screenshots/7.jpg)  

Result example in Excel  
![Screenshot](./screenshots/8.jpg)  

File structure for for screenshot example:  

│   searchinpdf.exe  
│  
└───test  
        matches.xlsx  
        tetst (1).pdf  
        tetst (10).pdf  
        tetst (11).pdf  
        tetst (12).pdf  
        tetst (2).pdf  
        tetst (3).pdf  
        tetst (4).pdf  
        tetst (5).pdf  
        tetst (6).pdf  
        tetst (7).pdf  
        tetst (8).pdf  
        tetst (9).pdf  


