Documentation for Aggregator Compiler Project

Purpose
Compile rates of various messaging routes from a variety of aggregators, across multiple file formats, currencies and other variables and generate a single file that easily displays all data.

Reference documentation : https://docs.google.com/document/d/11ky3JI2LIBvFbMy_wCNcmfo7TbuUMsz-_Ig7skWXSKU/edit?usp=sharing

Excel documents used for reference are in Compiled Data.  - DO NOT DELETE
Sheets for viewing are found in Rate Sheets.  
Relevant Files:
	Source_Compiler.py
	Database_Manipulation.py(c)
	CurrencyConverterNew.py(c)
	Google_API_Manipulation.py(c)
	Email_Notifications.py(c)
	write_log.py(c)

Packages needed:
apiclient==1.0.3
fixerio==0.1.1
google-api-python-client==1.6.2
gspread==0.6.2
httplib2==0.10.3
oauth2client==4.1.1
openpyxl==2.4.8
pdfminer==20140328
xlrd==1.0.0
xlutils==2.0.0
xlwt==1.2.0

