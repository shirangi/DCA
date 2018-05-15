# AUTHOR: Trent Bone
# DATE: 6/13/2017
# OBJECTIVE: Take long DI file and turn segmented csv portions into individual sheets by well
# FIXES: Need to figure out a way to wear the fluid type and API get drug down the whole way, but maybe we do not need it?

##### NECESSARY IMPORTS###############
from __future__ import division
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import easygui
from lmfit import*
from decimal import*
import os
import xlwt
from tqdm import tqdm
import xlrd
import xlsxwriter
import datetime
import pythoncom
import win32com.client as win32
import pandas as pd
############################
############################################
def DI_Downloads():
	############## OPENING FILE ##################
	File=easygui.fileopenbox(filetypes=['*.dri'])
	with open(File, 'r') as File:
		workbook=xlwt.Workbook()
		r=int()
		######## READING AND WRITING FILE###############
		for line in File:
			if "[B]" in line:
				fluid=line.find('GAS')
				if fluid !=-1:
					fluid=('GAS WELL')
				else:
					fluid=('OIL WELL')
				quotes=line.split('"')[1::2]

				lease=quotes[0]
				well_num=quotes[12]
				sheet_name=lease+well_num
				r=2
				n=1
				try:
					sheet=workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
				except Exception as e:
					sheet_name=sheet_name+'(1)'
			if "[C]" in line:
				commas=line.split(',')
				API=commas[1]
				sheet.write(0,1,"API")
				sheet.write(1,0,fluid)
				sheet.write(1,1,API)
			if "[D]" in line:
				commas=line.split(',')
				date=commas[1]
				oil=commas[2]
				gas=commas[3]
				water=commas[4]
				if oil=="":
					oil=0
				if gas=="":
					gas=0
				if water=="":
					water=0
				if oil=="nan":
					oil=0
				if gas=="nan":
					gas=0
				if water=="nan":
					water=0
				if float(oil)+float(water)+float(gas)>0:
					sheet.write(r,2,date)
					sheet.write(r,3, float(oil))
					sheet.write(r,4,float(gas))
					sheet.write(r,5,float(water))
				if r==2:
					first_date=datetime.datetime.strptime(date,'%m/%d/%Y')-datetime.timedelta(days=30)
					first_date=str(first_date.strftime('%m/%d/%Y'))
					sheet.write(1,2,first_date)
					sheet.write(1,3,0)
					sheet.write(1,4,0)
					sheet.write(1,5,0)
					sheet.write(0,0,'Well Type')
					sheet.write(0,2,"Date")
					sheet.write(0,3,"Oil (STB/Month)")
					sheet.write(0,4,"Gas (MSCF/Month)")
					sheet.write(0,5,"Water (STB/Month)")
				r=r+1
	File_Name=easygui.enterbox('Enter File Name to Save')
	workbook.save(File_Name+'.xls')
	########################################################

	########## GETTING PATH AND FILE########################
	script_dir=os.path.dirname(__file__)
	rel_path=str(File_Name+'.xls')
	File=os.path.join(script_dir,rel_path)
	#######################################################

	###### GETTING RID OF COPIES ###############
	workbook=xlrd.open_workbook(File)
	nsheets=len(workbook.sheets())
	sheet_names=workbook.sheet_names()
	###### MAKING LIST OF SHEETS TO DELETE ################
	ws_to_delete=[]
	sheet_index=[]
	for sheet in xrange(0,nsheets):
		ws=workbook.sheet_by_index(sheet)
		if ws.cell(0,0).value=='':
			index=sheet
			ws_to_delete.append(sheet_names[index])
		else:
			continue	
	############################################################	

	########### DELETING COPIES ################################
	pythoncom.CoInitialize()
	excel=win32.gencache.EnsureDispatch('Excel.Application')
	excel_file=excel.Workbooks.Open(File)
	#excel.Visible=True
	excel.Visible=False
	nsheets=len(ws_to_delete)
	for sheet in xrange(0,nsheets):
		delete_sheet=excel_file.Sheets(ws_to_delete[sheet])
		delete_sheet.Delete()
		excel_file.Save()
	excel_file.Close(True)
	##############################################################
	return File



