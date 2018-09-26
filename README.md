# Indepen-dance ITI3 Summer Internship 2018
ITI3 projects are summer internships offered to students of the School of Computing Science at the University of Glasgow. Students work for local charities (in pairs), helping them with their IT problems. Our goal was to create a system for a dance company to record client information and simplify class management. 

[ITI3](http://www.dcs.gla.ac.uk/~hcp/iti3/)
[Glasgow University Settlement)](http://gusettlement.org/)
 
## Goals of the project
* become more user friendly by not printing attendance registers
* streamline process of taking attendance in dance classes
* simplify data management
* simplify the way members are contacted when classes are cancelled (make a mailing list)

All code is available in [macros] folder.
A user is working with the system through master.xlsm 

## Flow

### Beginning of the term
1. In the beginning of the term admin uses the master spreadsheet to generate register files based on classes and members database sheets.
2. The Excel registers are converted using the conversion centre sheet to Google Sheet files.

### Weekly
1. Instructors access the attendance registers on their devices using Google Sheets App (online or offline)
2. Lessons take place. Instructors take attendance and collect payments for the class. Money collected is given to the admin.
3. Admin converts attendance registers back to Excel format. Using master spreadsheet, admin generates a weekly financial report, updates registers (type of payment might change, new people joining etc) and creates contact lists.
4. Excel registers are converted back to Google Sheets files, so instructors can continue using them on their devices.

### End of the term
1. Register files are archived in Excel format.

## Where to see our scripts?
All Excel based scripts are located in Class Information/[macros] folder.
Scripts for conversion between Excel and Google Sheets are located in Class Information/Conversion folder.

## Authors
* Marija Mumm, 2nd year Computing Science Student, University of Glasgow
* Sara Jakubiak, 2nd year Computing Science Student, University of Glasgow
