# TitanTools
TitanTools is an excel add-in that provides efficiency in your day to day tasks. 
From cleaning to exporting your data this utility has you covered with over 30 awesome time saving buttons at your fingertips AND 18 folder locations in a menu button for easy navigation on your PC.

## Compatibility: 
- PC Only
- Excel 2007 - 2016
- Works on 32-bit or 64-bit

## The TitanTools Project
This add-in project begain in 2016 as an off and on again hobby which now has evolved, been enhanced and tested rigorously.
Add-ins are a great addition to the Excel function suite that cut down time to perform your everyday tasks

>I always wanted something that was available online, free and easy to install at work without having to call IT as at work I don't have who have Admin installation priveledges. 
Through extensive research, learning, trial and error I have finally created an xlam Excel Add-in that anyone can download and install for Excel Windows easily and safely.

## Continual Improvement
Through feedback and my every day use I will continue to improve the functionality of this add-in. 
A change log will also be provided to provide information on updates.

## All Tool Descriptions
A full list of all the tool descriptions can downloaded as a pdf here: 
- [View online](https://github.com/ExcelTitan/TitanTools/blob/master/TitanTools_Descriptions.pdf)
- [Download](https://github.com/ExcelTitan/TitanTools/raw/master/TitanTools_Descriptions.pdf)

## Installation

- Unzip the zip file to extract the contents
- Move the add-in to your local Microsoft AddIns directory.
- Open Excel
- Click the File tab
- Click Options
- Click the Add-Ins category on the left bar of the Excel Options Dialog Box
- In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
- Click Browse
- Navigate to the xla or xlam add-in and select it.
- Make sure the checkbox next to the add-in is checked, then click OK.

If the add-in keeps disappearing when you restart Excel, follow the directions below.

Important Note: There are a few additional steps to this installation due to an Office Security Update released in July 2016.


## Trust the Folder Location

The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust 
the folder location are below.

-  Open the Excel Options menu
   - File > Options
   - Open the Trust Center menu and Add a new location
   - Trust Center > Trust Center Settings… > Trusted Locations > Add new location
   - Browse for the folder that you saved the add-in file in.
   - Press OK, the folder should now appear in the Trusted Locations list.
   - Press OK on the Trust Center and Excel Options menu to close the menus.

-  The add-in is now in a trusted location and the add-in will appear every time you open Excel.


## UNBLOCK THE FILE
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, 
you will need to Unblock the file by changing a file property.

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to 
repeat the steps above to unblock it.

Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then 
you should see the add-ins ribbon tab load every time you open Excel.

### Group 1: TitanTools
A group of miscellaneous utilities that I use all the time! 
The Save Backup utility is the best little tool I have made. It simply makes a copy of the active workbook and appends a timestamp to the file name.
The Range Info button provides information like Sum and Count which Excel already does...but this allows you to also copy the information.
A whole drop down menu dedicated to navigating your PC for special folder locations.
Another drop down menu for extra utilities that speed up those tasks you need to do now and then.

### Group 2: Trim and Clean Data
Dedicated to manipulating and cleaning your data you can also join and split text, add leading or trailing text to each cell and finally change case. 
When changing a sentence to Proper Case the following words will remain lower case if they are not the first word in the sentence:
- a, an, and, as, at, but, by, for, from, if, in, is,
- nor, of, off, on, or, so, the, till, to, up, yet

### Group 3: Sheet Tools
Delete all empty sheets in the active workbook, unhide even sheets very hidden with VBA.
An awesome tool is to create sheet names en masse using the selected cell names

### Group 4: Export Data
Choose from a variety of ways to export your data, whether whole sheets to a new workbook or the selected range to CSV or TXT formats.

### Group 5: Table Tools
A collection of useful tools to export the data in your tables at each unique value in the selected columns. Each value can be exported to a new sheet in the same or new workbook or create a new file from each value.
A quick sort button for when you need to index or get the original order of the table. 
A custom filter button will insert a column to the far right, placing an 'x' in each visible cell. Then rename the Column header to something meaningful.
