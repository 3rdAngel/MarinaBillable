# MarinaBillable

It is necessary to first create a folder called "YardCrew" in the "Documents" directory of your computer.
Add "ycxx.py", "custom.xls", "image.png", and "icon.ico" to this folder in order to run the program.
It is also necessary to have a program that can open .xlsx files, such as MS Office or LibreOffice.

The program has 3 windows. Upon starting, 2 windows are hidden, and the visible window has an image of 3 mega-yachts docked,
and above is a single dropdown menu labeled "Techs". Click "Techs" to activate the dropdown of the list of techs, chose one and 
another window will open. In the code, this is under the section '# Third Window' with the guizero label "Window3". 
This is the White Card that is to be filled out. All of the necessary data will appear in the smaller right side box to choose from. 
Intially, this box will be empty until a list of "Today's Boats" is created either by clicking the "Add Boats" button on the bottom right, 
or by using the "edit" tab in the menu bar of the original first window, which has only one option, "Today's List".
Either of these will open the last window, which in the code is under '# Second Window' and is the guizero window labeled simply "window".
This window has two list boxes. The left box is a complete list of current open work orders that contain operation codes 
(op codes) that apply to the Yard Crew.
The right list box is the list of boats that are being worked on today. 
Simply scroll and click on the left list to populate the right list. 
Clicking on the right list will delete the boat from "Today's List".
This list populates the list on the "White Card" window. 
By choosing a boat from the right listBox on the White Card, a list of op codes for that boat will then be shown to choose from.
After choosing an op code, then the key pad appears to enter the number of hours that were spent that day on the op code. 
The checkbox above defaults to "Billable" since most work is billable but is available to record the unusual case of "non-billable".
Also the key pad is missing a "0" and instead has a ".5" since hours are only rounded to the nearest half hour and
it is not possible to work more than 9 hours on an op code in one day.
There is a choice to "clear" the white card if a mistake is made, or to "close without saving".
Once a card entry is complete, the option buttons "next" and "send it!" become active.
"Next" allows for another entry to be added to this white card, since most days more than one op code will be worked on.
"Send it!" is the Yard Crew slang for finishing a job. The white card is complete and can be saved in the data csv file
which will then be used to populate the white card Excel file when the supervisor wants to print the daily report of all
of the combined white cards.
To open the excel file, click the "File" tab in the menu bar of the first window and chose a report. The report will be created and opened.
