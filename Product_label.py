from fpdf import Template 
import pandas as pd
import csv
from openpyxl import load_workbook
from tabulate import tabulate
from datetime import datetime
from fpdf import FPDF

# Read and store content of an excel file.
read_file = pd.read_excel (r"C:\Users\Rohit Tagala\Downloads\TestCustomer.Feb2023.xlsx")

# Write the dataframe object into csv file
read_file.to_csv ("BPR.csv",index = None,header=True)	
df = pd.DataFrame(pd.read_csv("BPR.csv"))
#print (df)

read_file = pd.read_csv (r'C:\Users\Rohit Tagala\OneDrive\Documents\GitHub\Python-Files-\BPR.csv')
read_file.to_excel (r'New_BPR.xlsx', index = None, header=True)
file = (r"C:\Users\Rohit Tagala\OneDrive\Documents\GitHub\Python-Files-\New_BPR.xlsx")

# Load the entire workbook.
wb = load_workbook(file)

# Load one worksheet.
ws = wb.active
name = ws["D329"].value
print (name)
daily_dose = ws["D335"].value
print (daily_dose)
caps_per_bottle = ws["D336"].value
print (caps_per_bottle)
lot_number = ws["H2"].value
mfg_date = ws["C5"].value
print (mfg_date)
customer_id = ws["F2"].value

## Making supplement chart 
item_name = []
percent = []
dosage = []
with open (r'C:\Users\Rohit Tagala\OneDrive\Documents\GitHub\Python-Files-\BPR.csv', 'r') as file:
    csv_reader = csv.reader(file)
    for row in csv_reader:
        item_name.append(row[2])
        percent.append(row[4])
        dosage.append(row[5])
df = pd.DataFrame({"Item_name":item_name[15:42],"percent":percent[15:42],"dosage":dosage[15:42]})
item = item_name[15:69]
perc = percent [15:69]
dosa = dosage[15:69]
header = ("Ingredient", "percent", "dosage")
Ingredient_table = (tabulate(df, headers = header, tablefmt = 'fancy_grid'))
#print (Ingredient_table)


elements = [{'name':'company_name','type': 'T','x1' : 4.0, 'y1': 30.0, 'x2':115.0,'y2':37.8, 'font': 'Helvetica', 'align': 'L', 'text': ''},
{'name':'caps_count','type': 'T','x1' : 4.0, 'y1': 35.0, 'x2':115.0,'y2':41, 'font': 'Helvetica', 'align': 'L', 'text': ''},
    { 'name': 'Name', 'type': 'T', 'x1': 118.0, 'y1': 46.0, 'x2': 135.0, 'y2': 25.0, 'font': 'Helvetica', 'bold': 1.0,'align': 'L', 'text': '', 'size':13},
    { 'name': 'P_S', 'type': 'T', 'x1': 115.0, 'y1': 49.0, 'x2': 132.0, 'y2': 30.0, 'font': 'Helvetica', 'bold': 0,'underline': 1.0 ,'align': 'C', 'text': '','size':8 },
    { 'name': 'Note:', 'type': 'T', 'x1': 4.0, 'y1': 50.0, 'x2': 75.0, 'y2': 45.0, 'font': 'Helvetica', 'bold': 1,'underline': 0 ,'align': 'L', 'text': '', 'size': 13 , 'multiline':False},
    { 'name': 'Note1', 'type': 'T', 'x1': 4.0, 'y1': 53.0, 'x2': 75.0, 'y2': 53.0, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6 , 'multiline':False},
    { 'name': 'Note2', 'type': 'T', 'x1': 4.0, 'y1': 56.0, 'x2': 75.0, 'y2': 56.0, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6 , 'multiline':False},
    { 'name': 'Note3', 'type': 'T', 'x1': 4.0, 'y1': 58.0, 'x2': 75.0, 'y2': 61.5, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6 , 'multiline':True},
    { 'name': 'Made', 'type': 'T', 'x1': 4.0, 'y1': 72.0, 'x2': 75.0, 'y2': 73, 'font': 'Helvetica', 'bold': 1,'underline': 0 ,'align': 'L', 'text': '', 'size': 9 , 'multiline':False},
    { 'name': 'Lot_number', 'type': 'T', 'x1': 4.0, 'y1': 77.0, 'x2': 75.0, 'y2': 76, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6 , 'multiline':False},
    { 'name': 'Mfg_date', 'type': 'T', 'x1': 4.0, 'y1': 80.0, 'x2': 75.0, 'y2': 81, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6 , 'multiline':False},
    { 'name': 'Customer_ID', 'type': 'T', 'x1': 4.0, 'y1': 84.0, 'x2': 75.0, 'y2': 85, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'Mfg_by', 'type': 'T', 'x1': 25.0, 'y1': 77.0, 'x2': 95.0, 'y2': 76, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'L', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'Dietary', 'type': 'T', 'x1': 110.0, 'y1': 77.0, 'x2': 150.0, 'y2': 77, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'C', 'text': '', 'size': 7.5, 'multiline':False},
    { 'name': 'caps_per_bottle', 'type': 'T', 'x1': 115.0, 'y1': 80.0, 'x2': 140.0, 'y2': 80, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'C', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'per_bottle', 'type': 'T', 'x1': 115.0, 'y1': 83.0, 'x2': 140.0, 'y2': 83, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'C', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'Containes : Soy', 'type': 'T', 'x1': 160.0, 'y1': 80.0, 'x2': 290.0, 'y2': 80, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'C', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'Other Ingredient: Microcrystalline Cellulose and vegetable capsules', 'type': 'T', 'x1': 178.0, 'y1': 83.0, 'x2': 320.0, 'y2': 83, 'font': 'Helvetica', 'bold': 0,'underline': 0 ,'align': 'C', 'text': '', 'size': 6, 'multiline':False},
    { 'name': 'box', 'type': 'B', 'x1': 15.0, 'y1': 15.0, 'x2': 185.0, 'y2': 260.0, 'font': 'Arial', 'size': 0.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 1, 'background': 1, 'align': 'I', 'Text' : None
   }]
f = Template(format='A4', elements=elements,title = 'Rohit', orientation= "landscape")
f.add_page()
f['company_name'] = "SUGGESTED DIRECTIONS:"
f['caps_count'] = ("Take " + daily_dose + " capsules daily with food")
f['Name'] = name
f['P_S'] = "Personalized Supplements"
f['Note:'] = "NOTE:"
f['Note1'] = "Do not use if safety-seal is broken or missing."
f['Note2'] = "Store tightly sealed in a cool, dry place."
f['Note3'] = "If you are pregnant, lactating or taking any medications,\nconsult with your primary health care practitioner prior to taking."
f['Made'] = "MADE BY: VITAMINLABS"
f['Lot_number'] = ("Lot#  " +   lot_number)
f['Mfg_date'] = ("Mfg. Date  " +   mfg_date)
f['Customer_ID'] = ("Customer ID: " +   customer_id)
f['Mfg_by'] = "Manufactured By :   Vitamin One Formulas LTD."
f['Dietary'] = "Dietary Supplements"
f['caps_per_bottle'] = (caps_per_bottle + " Capsules")
f['per_bottle'] = "Per Container"
f['Containes : Soy'] = "Containes : Soy"
f['Other Ingredient: Microcrystalline Cellulose and vegetable capsules'] = "Other Ingredient: Microcrystalline Cellulose and vegetable capsules"
f.render("temp.pdf") 

