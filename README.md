 from openpyxl import *
from Tkinter import *
import pandas as pd
import datetime



now = datetime.datetime.now()
nameday = now.strftime("%A")

 
wb = load_workbook('C:\Python27\userdata.xlsx')
wb2 = load_workbook('C:\Python27\userdata2.xlsx')  
  

sheet = wb.active 
nameday = wb2.active


  
  
def excel(): 
      
    
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
  
  
 
def focus1(event): 
     
    caseid.focus_set() 
  
  
 
def focus2(event): 
    
    casetitle.focus_set() 
  
  
 
def focus3(event): 
     
    cat.focus_set() 
  
  
 
def focus4(event): 
   
    sev.focus_set() 
  
  
 
def focus5(event): 
     
    opencase.focus_set() 
  
  
 
def focus6(event): 
   
    closecase.focus_set() 

def focus7(event): 
    
    stat.focus_set() 
  
  
 
def focus8(event): 
    
    ass.focus_set() 
  
  
 
def focus9(event): 
    
    esca.focus_set() 
  
  

def focus10(event): 
    
    timeresolve.focus_set() 
  
  
 
def focus11(event): 
     
    note.focus_set() 
  
 
def focus12(event): 
    
    padd.focus_set() 
  

def clear(): 
      
     
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
  
  

def insert(): 
      
    
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
  
        
        current_row = sheet.max_row 
        current_column = sheet.max_column

        current_row = nameday.max_row 
        current_column = nameday.max_column 
  
      
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
  
       
        wb.save('C:\Python27\userdata.xlsx')
        wb2.save('C:\Python27\userdata2.xlsx')
        
         
        caseid.focus_set() 
  
        
        clear() 
  
  
 
if __name__ == "__main__": 
      
    
    root = Tk() 
  
    
    root.configure(background='light green') 
  
    
    root.title("SNOC Cases") 
  
  
    root.geometry("600x600") 
  
    excel() 
  
     
    heading = Label(root, text="SNOC Cases ADIB Project", bg="light green") 
  
     
    case_id = Label(root, text="Case_ID", bg="light green") 
  
     
    case_title = Label(root, text="Case_Title", bg="light green") 
  
   
    category = Label(root, text="Category", bg="light green") 
  
    
    severity = Label(root, text="Severity", bg="light green") 
  
     
    opening_date = Label(root, text="Opening_Date", bg="light green") 
  
    
    closing_date = Label(root, text="Closing_Date", bg="light green") 
  
     
    status = Label(root, text="Status", bg="light green") 

    assignment = Label(root, text="Assignment", bg="light green") 

   
    escalation = Label(root, text="Escalation", bg="light green") 

   
    time_resolve = Label(root, text="Time_Resolve", bg="light green") 

    notes = Label(root, text="Notes", bg="light green") 

    padding = Label(root, text="Padding", bg="light green") 




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
  
    
    caseid.bind("<Return>", focus1) 
  
    
    casetitle.bind("<Return>", focus2) 
  
     
    cat.bind("<Return>", focus3) 
  
    
    sev.bind("<Return>", focus4) 
  
  
    opencase.bind("<Return>", focus5) 
  
     
    closecase.bind("<Return>", focus6) 

 
    stat.bind("<Return>", focus7) 
  
    
    ass.bind("<Return>", focus8) 
  
   
    esca.bind("<Return>", focus9) 
  
    
    timeresolve.bind("<Return>", focus10) 
  
    
    note.bind("<Return>", focus11) 

    
    padd.bind("<Return>", focus12)
  
    
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
  
   
    excel() 
  
   
    submit = Button(root, text="Submit", fg="Black", 
                            bg="Red", command=insert) 
    submit.grid(row=13, column=1) 
  
    
    root.mainloop() 
