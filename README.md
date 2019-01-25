# import openpyxl and tkinter modules 
from openpyxl import *
from Tkinter import *
import pandas as pd
import datetime
# globally declare wb and sheet variable 

#print the day of week
now = datetime.datetime.now()
nameday = now.strftime("%A")

# opening the existing excel file 
wb = load_workbook('C:\Python27\userdata.xlsx')
wb2 = load_workbook('C:\Python27\userdata2.xlsx')  
  
# create the sheet object 
sheet = wb.active 
nameday = wb2.active


  
  
def excel(): 
      
    # resize the width of columns in 
    # excel spreadsheet 
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 10
    sheet.column_dimensions['C'].width = 10
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20
    sheet.column_dimensions['F'].width = 40
    sheet.column_dimensions['G'].width = 50
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 10
    sheet.column_dimensions['C'].width = 10
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20

    nameday.column_dimensions['A'].width = 30
    nameday.column_dimensions['B'].width = 10
    nameday.column_dimensions['C'].width = 10
    nameday.column_dimensions['D'].width = 20
    nameday.column_dimensions['E'].width = 20
    nameday.column_dimensions['F'].width = 40
    nameday.column_dimensions['G'].width = 50
    nameday.column_dimensions['A'].width = 30
    nameday.column_dimensions['B'].width = 10
    nameday.column_dimensions['C'].width = 10
    nameday.column_dimensions['D'].width = 20
    nameday.column_dimensions['E'].width = 20
  
    # write given data to an excel spreadsheet 
    # at particular location 
    sheet.cell(row=1, column=1).value = "Case_ID"
    sheet.cell(row=1, column=2).value = "Case_Title"
    sheet.cell(row=1, column=3).value = "Category"
    sheet.cell(row=1, column=4).value = "Severity"
    sheet.cell(row=1, column=5).value = "Opening_Date"
    sheet.cell(row=1, column=6).value = "Closing_Date"
    sheet.cell(row=1, column=7).value = "Status"
    sheet.cell(row=1, column=8).value = "Assignment"
    sheet.cell(row=1, column=9).value = "Escalation"
    sheet.cell(row=1, column=10).value = "Time_Resolve"
    sheet.cell(row=1, column=11).value = "Notes"
    sheet.cell(row=1, column=12).value = "Padding"

    nameday.cell(row=1, column=1).value = "Case_ID"
    nameday.cell(row=1, column=2).value = "Case_Title"
    nameday.cell(row=1, column=3).value = "Category"
    nameday.cell(row=1, column=4).value = "Severity"
    nameday.cell(row=1, column=5).value = "Opening_Date"
    nameday.cell(row=1, column=6).value = "Closing_Date"
    nameday.cell(row=1, column=7).value = "Status"
    nameday.cell(row=1, column=8).value = "Assignment"
    nameday.cell(row=1, column=9).value = "Escalation"
    nameday.cell(row=1, column=10).value = "Time_Resolve"
    nameday.cell(row=1, column=11).value = "Notes"
    nameday.cell(row=1, column=12).value = "Padding"
  
  
# Function to set focus (cursor) 
def focus1(event): 
    # set focus on the course_field box 
    caseid.focus_set() 
  
  
# Function to set focus 
def focus2(event): 
    # set focus on the sem_field box 
    casetitle.focus_set() 
  
  
# Function to set focus 
def focus3(event): 
    # set focus on the form_no_field box 
    cat.focus_set() 
  
  
# Function to set focus 
def focus4(event): 
    # set focus on the contact_no_field box 
    sev.focus_set() 
  
  
# Function to set focus 
def focus5(event): 
    # set focus on the email_id_field box 
    opencase.focus_set() 
  
  
# Function to set focus 
def focus6(event): 
    # set focus on the address_field box 
    closecase.focus_set() 

def focus7(event): 
    # set focus on the course_field box 
    stat.focus_set() 
  
  
# Function to set focus 
def focus8(event): 
    # set focus on the sem_field box 
    ass.focus_set() 
  
  
# Function to set focus 
def focus9(event): 
    # set focus on the form_no_field box 
    esca.focus_set() 
  
  
# Function to set focus 
def focus10(event): 
    # set focus on the contact_no_field box 
    timeresolve.focus_set() 
  
  
# Function to set focus 
def focus11(event): 
    # set focus on the email_id_field box 
    note.focus_set() 
  
# Function to set focus 
def focus12(event): 
    # set focus on the email_id_field box 
    padd.focus_set() 
  
# Function for clearing the 
# contents of text entry boxes 
def clear(): 
      
    # clear the content of text entry box 
    caseid.delete(0, END) 
    casetitle.delete(0, END) 
    cat.delete(0, END) 
    sev.delete(0, END) 
    opencase.delete(0, END) 
    closecase.delete(0, END) 
    stat.delete(0, END) 
    ass.delete(0, END) 
    esca.delete(0, END) 
    timeresolve.delete(0, END) 
    note.delete(0, END) 
    padd.delete(0, END)
  
  
# Function to take data from GUI  
# window and write to an excel file 
def insert(): 
      
    # if user not fill any entry 
    # then print "empty input" 
    if (caseid.get() == "" and
        casetitle.get() == "" and
        cat.get() == "" and
        sev.get() == "" and
        opencase.get() == "" and
        closecase.get() == "" and
        stat.get() == "" and
        ass.get() == "" and
        esca.get() == "" and
        timeresolve.get() == "" and
        note.get() == "" and
        padd.get() == ""): 
              
        print("empty input") 
  
    else: 
  
        # assigning the max row and max column 
        # value upto which data is written 
        # in an excel sheet to the variable 
        current_row = sheet.max_row 
        current_column = sheet.max_column

        current_row = nameday.max_row 
        current_column = nameday.max_column 
  
        # get method returns current text 
        # as string which we write into 
        # excel spreadsheet at particular location 
        sheet.cell(row=current_row + 1, column=1).value = caseid.get() 
        sheet.cell(row=current_row + 1, column=2).value = casetitle.get() 
        sheet.cell(row=current_row + 1, column=3).value = cat.get() 
        sheet.cell(row=current_row + 1, column=4).value = sev.get() 
        sheet.cell(row=current_row + 1, column=5).value = opencase.get() 
        sheet.cell(row=current_row + 1, column=6).value = closecase.get() 
        sheet.cell(row=current_row + 1, column=7).value = stat.get() 
        sheet.cell(row=current_row + 1, column=8).value = ass.get() 
        sheet.cell(row=current_row + 1, column=9).value = esca.get() 
        sheet.cell(row=current_row + 1, column=10).value = timeresolve.get() 
        sheet.cell(row=current_row + 1, column=11).value = note.get() 
        sheet.cell(row=current_row + 1, column=12).value = padd.get() 

        nameday.cell(row=current_row + 1, column=1).value = caseid.get() 
        nameday.cell(row=current_row + 1, column=2).value = casetitle.get() 
        nameday.cell(row=current_row + 1, column=3).value = cat.get() 
        nameday.cell(row=current_row + 1, column=4).value = sev.get() 
        nameday.cell(row=current_row + 1, column=5).value = opencase.get() 
        nameday.cell(row=current_row + 1, column=6).value = closecase.get() 
        nameday.cell(row=current_row + 1, column=7).value = stat.get() 
        nameday.cell(row=current_row + 1, column=8).value = ass.get() 
        nameday.cell(row=current_row + 1, column=9).value = esca.get() 
        nameday.cell(row=current_row + 1, column=10).value = timeresolve.get() 
        nameday.cell(row=current_row + 1, column=11).value = note.get() 
        nameday.cell(row=current_row + 1, column=12).value = padd.get() 
  
        # save the file 
        wb.save('C:\Python27\userdata.xlsx')
        wb2.save('C:\Python27\userdata2.xlsx')
        
        # set focus on the name_field box 
        caseid.focus_set() 
  
        # call the clear() function 
        clear() 
  
  
# Driver code 
if __name__ == "__main__": 
      
    # create a GUI window 
    root = Tk() 
  
    # set the background colour of GUI window 
    root.configure(background='light green') 
  
    # set the title of GUI window 
    root.title("SNOC Cases") 
  
    # set the configuration of GUI window 
    root.geometry("600x600") 
  
    excel() 
  
    # create a Form label 
    heading = Label(root, text="SNOC Cases ADIB Project", bg="light green") 
  
    # create a Name label 
    case_id = Label(root, text="Case_ID", bg="light green") 
  
    # create a Course label 
    case_title = Label(root, text="Case_Title", bg="light green") 
  
    # create a Semester label 
    category = Label(root, text="Category", bg="light green") 
  
    # create a Form No. lable 
    severity = Label(root, text="Severity", bg="light green") 
  
    # create a Contact No. label 
    opening_date = Label(root, text="Opening_Date", bg="light green") 
  
    # create a Email id label 
    closing_date = Label(root, text="Closing_Date", bg="light green") 
  
    # create a address label 
    status = Label(root, text="Status", bg="light green") 

    # create a address label 
    assignment = Label(root, text="Assignment", bg="light green") 

    # create a address label 
    escalation = Label(root, text="Escalation", bg="light green") 

    # create a address label 
    time_resolve = Label(root, text="Time_Resolve", bg="light green") 

    notes = Label(root, text="Notes", bg="light green") 

    padding = Label(root, text="Padding", bg="light green") 



  
    # grid method is used for placing 
    # the widgets at respective positions 
    # in table like structure . 
    heading.grid(row=0, column=1) 
    case_id.grid(row=1, column=0) 
    case_title.grid(row=2, column=0) 
    category.grid(row=3, column=0) 
    severity.grid(row=4, column=0) 
    opening_date.grid(row=5, column=0) 
    closing_date.grid(row=6, column=0) 
    status.grid(row=7, column=0) 
    assignment.grid(row=8, column=0) 
    escalation.grid(row=9, column=0) 
    time_resolve.grid(row=10, column=0) 
    notes.grid(row=11, column=0) 
    padding.grid(row=12, column=0)
  
    # create a text entry box 
    # for typing the information 
    caseid = Entry(root) 
    casetitle = Entry(root) 
    cat = Entry(root) 
    sev = Entry(root) 
    opencase = Entry(root) 
    closecase = Entry(root) 
    stat = Entry(root)
    ass = Entry(root) 
    esca = Entry(root) 
    timeresolve = Entry(root) 
    note = Entry(root) 
    padd = Entry(root)
  
    # bind method of widget is used for 
    # the binding the function with the events 
  
    # whenever the enter key is pressed 
    # then call the focus1 function 
    caseid.bind("<Return>", focus1) 
  
    # whenever the enter key is pressed 
    # then call the focus2 function 
    casetitle.bind("<Return>", focus2) 
  
    # whenever the enter key is pressed 
    # then call the focus3 function 
    cat.bind("<Return>", focus3) 
  
    # whenever the enter key is pressed 
    # then call the focus4 function 
    sev.bind("<Return>", focus4) 
  
    # whenever the enter key is pressed 
    # then call the focus5 function 
    opencase.bind("<Return>", focus5) 
  
    # whenever the enter key is pressed 
    # then call the focus6 function 
    closecase.bind("<Return>", focus6) 

    # bind method of widget is used for 
    # the binding the function with the events 
  
    # whenever the enter key is pressed 
    # then call the focus1 function 
    stat.bind("<Return>", focus7) 
  
    # whenever the enter key is pressed 
    # then call the focus2 function 
    ass.bind("<Return>", focus8) 
  
    # whenever the enter key is pressed 
    # then call the focus3 function 
    esca.bind("<Return>", focus9) 
  
    # whenever the enter key is pressed 
    # then call the focus4 function 
    timeresolve.bind("<Return>", focus10) 
  
    # whenever the enter key is pressed 
    # then call the focus5 function 
    note.bind("<Return>", focus11) 

    # whenever the enter key is pressed 
    # then call the focus5 function 
    padd.bind("<Return>", focus12)
  
    # grid method is used for placing 
    # the widgets at respective positions 
    # in table like structure . 
    caseid.grid(row=1, column=1, ipadx="100") 
    casetitle.grid(row=2, column=1, ipadx="100") 
    cat.grid(row=3, column=1, ipadx="100") 
    sev.grid(row=4, column=1, ipadx="100") 
    opencase.grid(row=5, column=1, ipadx="100") 
    closecase.grid(row=6, column=1, ipadx="100") 
    stat.grid(row=7, column=1, ipadx="100") 
    ass.grid(row=8, column=1, ipadx="100") 
    esca.grid(row=9, column=1, ipadx="100") 
    timeresolve.grid(row=10, column=1, ipadx="100") 
    note.grid(row=11, column=1, ipadx="100") 
    padd.grid(row=12, column=1, ipadx="100") 
  
    # call excel function 
    excel() 
  
    # create a Submit Button and place into the root window 
    submit = Button(root, text="Submit", fg="Black", 
                            bg="Red", command=insert) 
    submit.grid(row=13, column=1) 
  
    # start the GUI 
    root.mainloop() 
