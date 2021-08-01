import win32com.client
import os
import pandas as pd


DATA_SHEET = r"D:\DhruvData\MIT Gaming\Certificates\Club Certs\Certificate Master List.xlsx"
TEMPLATE_LOC = r"D:\DhruvData\MIT Gaming\Certificates\Club Certs\Senior Cert Template.psd"
SAVE_LOC = r"D:\DhruvData\MIT Gaming\Certificates\Club Certs\Senior Certs"


# import data from spreadsheet
main_sheet  = pd.read_excel(DATA_SHEET, sheet_name = 'Club')

# create a data frame with only the relevant columns
name_list = pd.DataFrame(main_sheet, columns=['Name:', 'Serial Number:'])

# get total number of entries in the sheet
total_names = name_list.shape[0]

# Open Photoshop Application
psApp = win32com.client.Dispatch("Photoshop.Application")

# Open Template File
psApp.Open(TEMPLATE_LOC)

# Set doc to refer to the open Active Document
doc = psApp.Application.ActiveDocument

# Assign variables to refer to relevant layers within the document
layer_name = doc.ArtLayers["Name"]
layer_sn = doc.ArtLayers["Serial"]

# Assign a variable referring to the text item within those layers
name_text = layer_name.TextItem
sn_text = layer_sn.TextItem

# create a loop to change names and save certificate as pdf
for i in range (total_names):
    # extract name and serial number from data frame
    name = str(name_list.iloc[i,0])
    ser_num = str(name_list.iloc[i,1])
    
    # change text contents of respective layers in template
    name_text.contents = name
    sn_text.contents = "Sr. No.: " + ser_num
    
    # set the save name and save location. The "/" in the serial number is replaced with "_" to adhere to file save name convention. 
    save_name = ser_num.replace('/', '_') + ".pdf"
    loc = SAVE_LOC + "\\" + save_name

    # Set pdf save options.
    saveOptions = win32com.client.Dispatch('Photoshop.PDFSaveOptions')
    
    saveOptions.Encoding = 2
    saveOptions.DownSample = 3
    saveOptions.DownSampleSize = 150
    saveOptions.DownSampleSizeLimit = 225
    saveOptions.JPEGQuality = 12
    saveOptions.PDFCompatibility = 3
    saveOptions.OptimizeForWeb = True
    saveOptions.PreserveEditing = False
    
    # saveOptions.presetFile = "PDF2"
    
    # document is saved as a copy
    doc.SaveAs(loc, saveOptions, True)

# psApp.Close(1)