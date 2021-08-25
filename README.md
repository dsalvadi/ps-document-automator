# Photoshop Document Automator
This is a python script that allows for the automated batch creation of documents in photoshop, such as certificates.

To use this script, you'll need to install the following packages:
```
pip install pypiwin32
pip install pandas
```
You will also require Photoshop installed on your system.

This script opens a template file in Photoshop and changes text details in specific fields (Name and Serial No). These fields are pre-named and formatted in the template file. It then proceeds to save the document as a PDF file, with the serial number as the file name. The encoding and compression for the PDF file is configured. This script loops through entries in a spreadsheet to process and create automated certificates for all entries. 
