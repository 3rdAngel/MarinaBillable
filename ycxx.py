# Imports
from guizero import App, Box, Text, ListBox, Picture, PushButton, info, Combo, MenuBar, Window, CheckBox
import webbrowser
import csv
from pathlib import Path
from datetime import date
from pandas import * #this allows calling ExcelFile
import xlsxwriter

# PATH TO DIRECTORY OF PROGRAM FILES
path = Path.cwd()
main_image = path.joinpath('image.png')  # This image must be in the same directory as the program
wcDir = Path.home() / 'Documents' / 'WhiteCards'  # The WhiteCards folder is in 'Documents' for easy access
wcDir.mkdir(exist_ok=True) # if folder doesn't exist, make directory
wc_folder = str(wcDir)

# Get current date
today = date.today()
date = today.strftime("%m/%d/%y") 

# DATED FILE NAMES TO OPEN:
dailyDataCsv = today.strftime("%y-%m-%d-data.csv") # file to save today's work hours from white card
yearlyDataCsv = today.strftime("%Y-data.csv") # file to add today's data into yearly totals
datedFile = wc_folder + today.strftime("/%y-%m-%d-whiteCard.xlsx") # Name for daily white card
yearDateFile = wc_folder + today.strftime("/%Y-whiteCard.xlsx") # Name for current year report of white card totals

# IF FILE DOESN'T EXIST, CREATE IT:
try:
    ddc = open(dailyDataCsv)
except:
    ddc = open(dailyDataCsv, 'w')
finally:
    ddc.close()
    
try:
    ydc = open(yearlyDataCsv)
except:
    ydc = open(yearlyDataCsv, 'w')
finally:
    ddc.close()

try:
    nl = open('NameList.txt')
except:
    nl = open('NameList.txt', 'w')
finally:
    nl.close()

try:
    oS = open('opSearch.txt')
except:
    oS = open('opSearch.txt', 'w')
    oS.write('SHRINK\nHULL ONLY\nQUICK\nPAINT\nANTIFOUL\nWAX\nW/C/W')
finally:
    oS.close()

try:
    TL = open('TechList.txt')
except:
    TL = open('TechList.txt', 'w')
    TL.write('20 Caleb\n28 Chad\n36 Joe\n57 Josh\n61 Cam\n68 Ralphy\n88 Odvan\n95 Kathryn\n103 Zach')
finally:
    TL.close()

try:
    ch = open('custom_headers.txt')
except:
    ch = open('custom_headers.txt', 'w')
    ch.write('Id\nName\nBoat\nMake\nModel\nOp\nDescript\nDate')
finally:
    ch.close()    
    
# INITIALIZED DATA STRUCTURES
boats = dict()
boatlist = []
todays_boats = [] #"TODAY'S BOATS"
opsForBoat = ["OpCodes"]

# List of words to search in the op code description to gather only the op codes related to the Yard Crew
ops = []
with open('opSearch.txt', 'r') as opf:
    for opterm in opf:
        term = opterm.strip()
        ops.append(term)

# Get Tech Names and load dictionary for the header:
techs = {}
techList = []
techCombo = []
with open('TechList.txt', 'r') as techf:
    for tec in techf:
        tec = tec.strip()
        t = tec.split()
        techs[t[0]] = t[1]
        techList.append(t[0])
        techCombo.append(t[1])

# LOAD LIST OF COLUMN HEADERS FOR 'custom.xls' FILE BASED ON QUERY FIELDS IN REPORT
col_hds = []
with open('custom_headers.txt', 'r') as chf: #chf=custom headers file
    for label in chf:
        label = label.strip()
        col_hds.append(label)

#FUNCTION:
    
def alphab(boatsVal):                         #alphabetize the boat dictionary by names
    return boatsVal[1].get('Name')

#DATA MANIPULATION:

xls = ExcelFile('custom.xls')                 #Convert .xls file to a dictionary (report)

df = xls.parse(sheet_names=0, header=None, names=col_hds)

new = df.to_dict(orient='records')                   #dataframe to dictionary, orient records makes each line a dict
                                                     #where each header is a key matched to a value in the excel row
new = [x for x in new if isinstance(x['Name'], str)] #rewrites dict with only entries where Name is a string

for record in new:
    nm = record['Name'].split(",", 1)                #Split 'Name' only once at the first ","
    record['Name'] = nm[0]                           #reduces the Name to last name only
    for keyy in record:
        if isinstance(record[keyy], float):          #gets rid of "nan" entries
            record[keyy] = ""
    
    for opc in ops: #for op code in op codes keyword list (ops) - test each line of dictionary for relavant opcodes
        if isinstance(record['Descript'], str) and record['Descript'].find(opc) >= 0: #if .find()method returns 0, there is a match. -1 is not a match
            if record['Id'] in boats.keys(): #ln['Id'] is the work order number of the line, boats.keys are the wo numbers already in the dictionary of relevant boats
                boats[record['Id']]['Codes'][record['Op']] = record['Descript'] #if it's already there, just add more op codes
            else:                                                  #if it's not there, add all details
                boats[record['Id']] = {'Name' : nm[0], 'Boat' : record['Boat'],'Make' : record['Make'],\
                                        'Model' : record['Model'], 'Codes' : {record['Op'] : record['Descript']}}
                
alphaboats = sorted(boats.items(), key=alphab) #sort boats{} into an alphabetized list of dictionaries called alphaboats
for entry in alphaboats: #creates boatlist to display in the "All Work Orders" ListBox
    boatlist.append('{} {} "{}" {} {}'.format(entry[0], entry[1]['Name'], entry[1]['Boat'],\
                                              entry[1]['Make'], entry[1]['Model']))

# populate "TODAY'S BOATS" and white card list
with open('NameList.txt', 'r') as NL:
    for nam in NL:
        nam = nam.strip()
        x = boats.get(nam)
        xy = '{} "{}" {}'.format(x['Name'], x['Boat'], nam)
        todays_boats.append(xy)

# FUNCTIONS

def todays_report():
    create_xlsx(datedFile, dailyDataCsv)
    
def yearly_report():
    create_xlsx(yearDateFile, yearlyDataCsv)
    
def white_card_folder():
    webbrowser.open(wc_folder)

def open_window():
    window.show()
    
def close_window():
    window.hide()

def close_white_card():
    window3.hide()
    cardLabel.clear()
    card.clear()
    opListBox.hide()
    opListBox.clear()
    billable_checkBox.hide()
    keypad.hide()
        

def info_button():
    info("Info", 'If the name or work order number you are looking for is not listed, try updating the "Custom" report from DockMaster. Be sure to save it as Custom.xls in the YardCrew folder')

def display_wo(value): #displays the workorders that were written to the output file
    display.append(value)
        
def create_xlsx(filename, csvDataSource):

# Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
    Bold_Center = workbook.add_format({'bold': True, 'center_across': True})
    Centered = workbook.add_format({'center_across': True})
 
# Adjust the column width.
    worksheet.set_column('A:A', 10)
    worksheet.set_column('C:C', 20, Centered)
    worksheet.set_column('D:D', 4, Centered)
    worksheet.set_column('E:N', None, Centered)

# Write today's date in the first cell
    worksheet.write('A1', date)

# Write some data headers.
    for i in range(len(techs)):
        worksheet.write(0, i + 4, techs[techList[i]], Bold_Center)
    for i in range(len(techs)):
        worksheet.write(1, i + 4, techList[i], Bold_Center)

    worksheet.write('A2', 'RO#', Bold_Center)
    worksheet.write('B2', 'OP#', Bold_Center)
    worksheet.write('C2', 'CUSTOMER', Bold_Center)
    worksheet.write('D2', 'Bill', Bold_Center)

# Some data we want to write to the worksheet.
    work = []                                 # to read in csv data
    with open(csvDataSource, 'r') as csvf:
        csvreader = csv.reader(csvf)
        for line in csvreader:
            if len(line) > 0:
                work.append(line)

# Start from the first cell below the headers.
    row = 2
    col = 0

    for d in (work):  # d is data line in the work list from the csv data
        l = len(d)    # to find number of hour entries, length is number of items in data line 
        worksheet.write_string(row, col, d[0])
        worksheet.write_string(row, col + 1, d[1])
        worksheet.write_string(row, col + 2, d[2])
        worksheet.write_string(row, col + 3, d[3])
        for j in range(4, l, 2):
            tID = techList.index(d[j]) # find the index number of the tech in the techList, this matches the placement on the white card xls
            hrs = d[j+1]
            worksheet.write_string(row, tID + 4, hrs)
        row += 1

    workbook.close()
    webbrowser.open(filename)

def write_wo(infoStr):
    infoList = infoStr.split()

    nmbt = '{} "{}" {}'.format(infoList[1], infoList[2], infoList[0]) #displays Name Boat Name and Work Order ID
    display.append(nmbt)
    boat.append(nmbt)
    with open("NameList.txt", "a") as nl:
        nl.write('{}\n'.format(infoList[0]))
    global prntr
    prntr = 0

def delete_name(nm):
    display.remove(nm) #remove name from list on screen
    global boats 
    i = 0
    split_nm = nm.split()
    ID = split_nm[-1]  
    
    with open('NameList.txt', 'r') as old: #read in old list of names
        ids = old.readlines()
    with open('NameList.txt', 'w') as new: #write updated list
        for Id in ids: #interate line by line, each line is a name
            if Id.strip() != ID: #do not write deleted name
                new.write(Id) #write the new list without the deleted name

def writeTech(techName):
    if techName != "Tech":
        cardLabel.clear()
        cardLabel.append('{}{}\t{}'.format("Tech: ", techName, date))
        window3.show()
        boat.show()
        add_boat.show()
        Next.disable()
        done.disable()
        Next.bg="light grey"
        done.bg="light grey"

#todo
def select_boat(selected):
    selectList = selected.split()
    selectName = selectList[0]
    selectWO = selectList[-1]
    global boats
    for k, v in boats.items(): # key, value
        if k == selectWO:
            card.append(k)
            for o, d in v['Codes'].items():  # opcode, description in value[]  # can also be- for oc in v['Codes']:
                opListBox.append("{}  {}".format(d, o))   # can also be-   (oc, v['Codes'][oc])
    
    card.append("\t")
    card.append(selectName)
    sp = 10 - len(selectName) # spacing for next item
    for j in range(0, sp, 1):
        card.append(" ")
    boat.hide()
    add_boat.hide()
    opListBox.show()
    close_whiteCard.show()
                
def select_op_code(description):
    if description != "OpCodes":
        code = description.split()
        card.append(code[-1])
        if len(code[-1]) == 7:
            card.append("  hours: ")
        else:
            card.append("\t  hours: ")
        opListBox.hide()
        billable_checkBox.show()
        billable_checkBox.value=1
        keypad.show()
        opListBox.clear()

def hours(add):
    card.append(add)
    button1.disable()
    button2.disable()
    button3.disable()
    button4.disable()
    button5.disable()
    button6.disable()
    button7.disable()
    button8.disable()
    button9.disable()
    Next.bg="yellow"
    done.bg="light green"
    Next.enable()
    done.enable()
    
def half(add):
    card.append(add)
    button1.disable()
    button2.disable()
    button3.disable()
    button4.disable()
    button5.disable()
    button6.disable()
    button7.disable()
    button8.disable()
    button9.disable()
    button0.disable()
    Next.bg="yellow"
    done.bg="light green"
    Next.enable()
    done.enable()
    
def next_line():
    card.append("\n") #ready to start the next line on the white card
    billable_checkBox.hide()
    keypad.hide()
    boat.show()
    button1.enable()
    button2.enable()
    button3.enable()
    button4.enable()
    button5.enable()
    button6.enable()
    button7.enable()
    button8.enable()
    button9.enable()
    button0.enable()
    Next.disable()
    done.disable()
    Next.bg="light grey"
    done.bg="light grey"

def print_white_card():
    window3.hide()
    
    if billable_checkBox.value == 1:
        bill = "B"
    else:
        bill = "NB"
    
    cl = (cardLabel.value).split()
    c = (card.value).split()
    for item in techs:
        if techs[item] == cl[1]:
            ID = item

    with open("WhiteCardLog.txt", "a") as wc:
        wc.write("{}  {} {} {}  {} {} {} {}\n".format(cl[2], ID, cl[1], c[0], c[2], c[1], bill, c[4]))
    new_entry = [c[0], c[2], c[1], bill, ID, c[4]]

# APPEND YEARLY DATA CSV
    original = []
    match = 0
    
# OPEN YEARLY DATA CSV LOOK FOR A MATCH, APPEND DATA, SAVE AS 'ORIGINAL'
    with open(yearlyDataCsv, "r") as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            if len(row) > 0:
                if row[0] == new_entry[0] and row[1] == new_entry[1] and row[3] == new_entry[3]:
                    row.append(new_entry[4])
                    row.append(new_entry[5])
                    match += 1
                original.append(row)

# IF NEW ENTRY BOAT IS NOT ALREADY ON LIST, JUST APPEND A NEW ROW
    if match == 0:
        with open(yearlyDataCsv, "a") as apd:
            csvappend = csv.writer(apd)
            csvappend.writerow(new_entry)
# IF NEW ENTRY BOAT IS ALREADY ON LIST, WRITE OVER OLD CSV WITH CONTENT IN 'ORIGINAL'
    else:
        with open(yearlyDataCsv, "w") as wtr:
            csvwriter = csv.writer(wtr)
            for oldRow in original:
                csvwriter.writerow(oldRow)
                
# APPEND DAILY DATA CSV
    original = []
    match = 0
    
# OPEN DAILY CSV LOOK FOR A MATCH, APPEND DATA, SAVE AS 'ORIGINAL'
    with open(dailyDataCsv, "r") as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            if len(row) > 0:
                if row[0] == new_entry[0] and row[1] == new_entry[1] and row[3] == new_entry[3]:
                    row.append(new_entry[4])
                    row.append(new_entry[5])
                    match += 1
                original.append(row)

# IF NEW ENTRY BOAT IS NOT ALREADY ON LIST, JUST APPEND A NEW ROW
    if match == 0:
        with open(dailyDataCsv, "a") as apd:
            csvappend = csv.writer(apd)
            csvappend.writerow(new_entry)
# IF NEW ENTRY BOAT IS ALREADY ON LIST, WRITE OVER OLD CSV WITH CONTENT IN 'ORIGINAL'
    else:
        with open(dailyDataCsv, "w") as wtr:
            csvwriter = csv.writer(wtr)
            for oldRow in original:
                csvwriter.writerow(oldRow)
                
    tech.select_default()
    cardLabel.clear()
    card.clear()
    billable_checkBox.hide()
    keypad.hide()
    button1.enable()
    button2.enable()
    button3.enable()
    button4.enable()
    button5.enable()
    button6.enable()
    button7.enable()
    button8.enable()
    button9.enable()
    button0.enable()
    
def clear_button():
    card.clear()
    boat.show()
    add_boat.show()
    opListBox.clear()
    opListBox.hide()
    billable_checkBox.hide()
    keypad.hide()
    button1.enable()
    button2.enable()
    button3.enable()
    button4.enable()
    button5.enable()
    button6.enable()
    button7.enable()
    button8.enable()
    button9.enable()
    button0.enable()
    

def close_app():
    app.destroy()

# App
app = App("Hyannis Marina", width=1140, height=700)
app.bg = (115, 160, 230)
#app.tk.attributes("-fullscreen",True)

# Widgets
menubar = MenuBar(app,
                  toplevel=["File", "Edit"],
                  options=[
                      [ ["Today's Report", todays_report], [today.strftime("%Y Report"), yearly_report],\
                        ["Recent Reports", white_card_folder]],
                      [ ["Today's List", open_window]]
                  ])

title = Text(app, "Yard Crew", color="White", size=30, font="Quicksand Medium")
tech = Combo(app, options=techCombo + ["Tech"], selected="Tech" , command=writeTech)
tech.text_size = 30
tech.bg = "white"
spacer = Box(app, width="fill", height=20)
pic = Picture(app, image=str(main_image))
pic.width = 460
pic.height = 345
spacer2 = Box(app, width="fill", height=30)
close = PushButton(app, text="Exit", command=close_app)

# Second Window: Edit Today's List
window = Window(app, title = "Edit Today's List", height=500, width=950, layout="grid")
window.hide()

lspacer = Box(window, grid=[0,0], width=20, height=500)
left_box = Box(window, grid=[1,0])
mspacer = Box(window, grid=[2,0], width=40, height=500)
right_box = Box(window, grid=[3,0])
rspacer = Box(window, grid=[4,0], width=20, height=500)

# LeftBox
listBoxTitle = Text(left_box, "All Work Orders", grid=[0,0], size=16, color="Dark Blue",\
                       font="Quicksand Medium")
instructions = Text(left_box, "click to add boat to Today's List", grid=[0,1], color="Dark Blue")
fullList = ListBox(left_box, items = boatlist, command = write_wo, width = 500, height = 345, grid=[0,2],\
                      scrollbar=True)  
fullList.bg = (200, 230, 255)
fullList.text_size = 16
fullList.tk.children["!scrollbar"].config(width=25) # access the tkinter object to resize the scrollbar
fullList._listbox.resize(None, None) # to show scrollbar, reset internal listbox size
fullList._listbox.resize("fill", "fill") # now fill available space

# Missing Name Button
infobutton = PushButton(left_box, text = "Missing Name?", padx=5, command = info_button, align="bottom", grid=[0,3])
infobutton.text_color = "Dark Blue"
infobutton.text_size = 14
infobutton.bg = (110, 170, 255)

# RightBox
todaysList = Text(right_box, text="Today's List", size=16, color="Dark Blue", font="Quicksand Medium")
listInstruct = Text(right_box, text="click to remove boat from Today's List", color="Dark Blue")
display = ListBox(right_box, items = todays_boats, command = delete_name ,\
                  width = 350, height = 345, scrollbar=True)
display.bg = (200, 230, 255)
display.text_size=16
display.tk.children["!scrollbar"].config(width=25) # access the tkinter object to resize the scrollbar
display._listbox.resize(None, None) # to show scrollbar, reset internal listbox size
display._listbox.resize("fill", "fill") # now fill available space

finished = PushButton(right_box, text = "Finished", padx=50, command = close_window, align="bottom")
finished.text_color = "Dark Blue"
finished.text_size=14
finished.bg = (110, 170, 255)

# Third Window
window3 = Window(app, title = "White Card", height=500, width=950, layout="grid")
window3.hide()
lt_box = Box(window3, grid=[0,0], width=650, height=500, border=True)
lt_box.bg="white"
rt_box = Box(window3, grid=[1,0], width=300, height=500, border=True)


boat = ListBox(rt_box, items=todays_boats, command=select_boat, width=280, height=450)
boat.text_size=16
opListBox = ListBox(rt_box, items=opsForBoat, command=select_op_code, visible=False, width=280, height=450)
opListBox.text_size=16
add_boat = PushButton(rt_box, text = "Add Boat", padx=50, command = open_window, align="bottom")
add_boat.text_color = "Dark Blue"
add_boat.text_size=14
add_boat.bg = (110, 170, 255)

close_whiteCard = PushButton(rt_box, text = "Close without Saving", command = close_white_card , visible=False, align="bottom")
close_whiteCard.text_size=14

billable_checkBox = CheckBox(rt_box, text="BILLABLE HOURS", visible=False)
billable_checkBox.text_size=12
keypad = Box(rt_box, layout="grid", visible=False)
button1 = PushButton(keypad, command=hours, args=["1"], padx=20, text="1", grid=[0,0])
button2 = PushButton(keypad, command=hours, args=["2"], padx=20, text="2", grid=[1,0])
button3  = PushButton(keypad, command=hours, args=["3"], padx=20, text="3", grid=[2,0])
button4  = PushButton(keypad, command=hours, args=["4"], padx=20, text="4", grid=[0,1])
button5  = PushButton(keypad, command=hours, args=["5"], padx=20, text="5", grid=[1,1])
button6  = PushButton(keypad, command=hours, args=["6"], padx=20, text="6", grid=[2,1])
button7  = PushButton(keypad, command=hours, args=["7"], padx=20, text="7", grid=[0,2])
button8  = PushButton(keypad, command=hours, args=["8"], padx=20, text="8", grid=[1,2])
button9  = PushButton(keypad, command=hours, args=["9"], padx=20, text="9", grid=[2,2])
button0  = PushButton(keypad, command=half, args=[".5"], padx=12, text=".5", grid=[1,3])
button1.text_size=30
button2.text_size=30
button3.text_size=30
button4.text_size=30
button5.text_size=30
button6.text_size=30
button7.text_size=30
button8.text_size=30
button9.text_size=30
button0.text_size=30
button1.bg="white"
button2.bg="white"
button3.bg="white"
button4.bg="white"
button5.bg="white"
button6.bg="white"
button7.bg="white"
button8.bg="white"
button9.bg="white"
button0.bg="white" #(200, 230, 250)

cardLabel = Text(lt_box, text="Tech: ", size=12, width=200)
card = Text(lt_box, text="",size=20, width=200)
card.tk.config(justify='left')
buttons = Box(lt_box, align="bottom", width=502, height=50, layout="grid")
clear = PushButton(buttons, command=clear_button, text="Clear", padx=20, grid=[0,0])
clear.bg=(250, 200, 200)
left_gap = Box(buttons, grid=[1,0], width=50, height=50)
Next = PushButton(buttons, command=next_line, text="Next", padx=100, grid=[2,0])
Next.bg="light grey"
right_gap = Box(buttons, grid=[3,0], width=50, height=50)
done = PushButton(buttons, command=print_white_card, text="Send it!", padx=15, grid=[4,0])
done.bg="light grey"

# Display
app.display()
