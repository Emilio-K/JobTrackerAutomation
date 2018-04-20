from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf
from datetime import date
import os, time, sys

cwd= os.getcwd()
path_to_watch=cwd + "\Job Applications\\2018"
workbook_path = cwd + "\jobs.xlsx"
today= str(date.today())
START_ROW=47
rb=open_workbook(workbook_path)
r_sheet = rb.sheet_by_index(0)
wb=copy(rb)
w_sheet=wb.get_sheet(0)
col_date=0
col_company = 1
col_contacted=2
col_jobposting=3

def writeEntry(added):
		row_index = START_ROW;
		w_sheet.write(row_index, col_date,today)
		w_sheet.write(row_index, col_company, str(added))
		w_sheet.write(row_index,col_contacted, 'No')
		curr_job_posting = "file:\\\\" + cwd + "\Job Applications\\2018\\" + added + "\\jobposting.docx"
		w_sheet.write(row_index, col_jobposting, curr_job_posting)
		wb.save('jobs.xls')
		++START_ROW;

#create initial list of directories in the path.
before = dict([(f,None) for f in os.listdir (path_to_watch)])
while True:
	time.sleep(60*60)
	
	#create list after to compare with the before list
	after=dict([(f,None) for f in os.listdir (path_to_watch)])
	#create list of files/folders that were added
	added = [f for f in after if not f in before]
	#loop through those  files and write enteries to excel sheet.
	for i in added:
	 	writeEntry(i)
	before = after
	#if added, then we want to update the excel worksheet with the folder name


	

	


