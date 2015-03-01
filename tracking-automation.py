#Stage 2: Open the External Machine
machine = "C:\\Users\\A_DO\\Dropbox\\1. The Machine\\VOD (Dom & UK) rev70.xlsx"
#This is where u tell python what u want to label your specific excel file and where it's at
from win32com.client import Dispatch
excel = Dispatch("Excel.Application")
excel.Visible = 1

excel.Workbooks.Open (machine)
excel.Worksheets("Tracker (domestic) rep").Select() #this is where you specify the tab you want to go to
excel.Range("A2:AR30000").Select() #This is where u specify the range u want to select
excel.Selection.ClearContents() 

excel.Worksheets("Tracker (UK) rep").Select() 
excel.Range("A2:AR30000").Select() 
excel.Selection.ClearContents() 

excel.Worksheets("QC3 log (from M) rep").Select()
excel.Range("A2:AR30000").Select()
excel.Selection.ClearContents()

excel.Worksheets("CMC Portal report (CSV) rep").Select()
excel.Range("A2:AR30000").Select()
excel.Selection.ClearContents()

excel.Worksheets("EDM QC3 Pending (daily) rep").Select()
excel.Range("A2:AR30000").Select()
excel.Selection.ClearContents()

XlDirectionDown = 4
excel.Worksheets("CMC Aspera (daily email) app").Select()
excel.Range("A:A").End(XlDirectionDown).Select() #this goes to the end of the range you specify.
excel.ActiveCell.Offset(2,1).Select() #this goes one range down to the range you specified above