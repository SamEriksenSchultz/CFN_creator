#spreadsheet library
import xlrd
import PySimpleGUI as sg

layout = [[sg.T("")], [sg.Text("Select UCR request form:")], [sg.Input(key="file_path_input"), sg.FileBrowse()], 
         [sg.Text("Enter price:")], [sg.InputText(key="price")], 
         [sg.Text("Enter file destination folder:")], [sg.Input(key="-IN2-" ,change_submits=True), sg.FolderBrowse(key="destination_folder")],
         [sg.Button("Confirm")]]

#create the Window
window = sg.Window('CFN creator 1.0', layout, size = (425,300))

#default file_path, price, and destination_folder values
file_path = "null"
price = "0"
destination_folder = "null"

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Confirm":
        
        #gets file path from user input
        file_path = values["file_path_input"]
        
        #gets price from user input
        price = values["price"]
        
        #gets destination folder form user input
        destination_folder = values["destination_folder"]
        
        window.close()
        

wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

#getting column with UCR-IDs 
UCR_ID_col = 0
for i in range(sheet.ncols):
  if sheet.cell_value(0, i) == "UCR ID":
    UCR_ID_col = i
    
#gets UCR-IDs from spreadsheet    
for i in range(sheet.nrows-1):
  
  UCR_ID = sheet.cell_value(i + 1, UCR_ID_col)
  
  #corrects UCR-ID naming for Windows file names
  file_name = UCR_ID.replace('/', '-')
  
  #creates CFN file 
  f = open(destination_folder + "\\" + file_name + ".CFN", 'x')
  f.write("SMARTECARTE_ENABLED=1\nSMARTECARTE_DEFAULT_AMOUNT="+price+"\nSMARTECARTE_UCR_ID="+UCR_ID)
  
  
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Confirm":
        
        #gets file path from user input
        file_path = values["file_path_input"]
        
        #gets price from user input
        price = values["price"]
        
        #gets destination folder form user input
        destination_folder = values["destination_folder"]
        
        window.close()
         
window.close()