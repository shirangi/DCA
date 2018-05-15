# Author: Trent Bone
# Date: 2/14/2018

# NECESSARY IMPORTS
###########################
from __future__ import division
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import easygui
import xlrd
from lmfit import*
import math
from scipy import stats
import pandas as pd
from datetime import datetime
from lmfit import*
from decimal import*
import os
import xlwt
from tqdm import tqdm
import win32com.client as win32
############################
###########################

# Opens file
def open_file(File):
	#File=easygui.fileopenbox()
	workbook=xlrd.open_workbook(File)
	number_of_sheets=workbook.nsheets 
	return number_of_sheets


# Puts it into pandas
def DRI_to_Dataframe(File,nsheets):
	wb=xlrd.open_workbook(File)
	number_of_sheets=nsheets
	production_df={}
	for sheet in xrange(0,number_of_sheets):
		production_df[sheet]=pd.read_excel(File,sheetname=sheet)
	for frame in xrange(0,number_of_sheets):
		production_df[frame].fillna(0, inplace=True)
		Dates=production_df[frame].loc[:,"Date"]
		shape=production_df[frame].shape
		months=shape[0]
		time=[30]
		for month in xrange(1,months):
			this_month=Dates[month]
			last_month=Dates[month-1]
			date_format="%m/%d/%Y"
			delta_t=datetime.strptime(str(this_month),date_format)-datetime.strptime(str(last_month),date_format)
			delta_t=delta_t.days
			time.append(delta_t)
		production_df[frame]['Delta_t (Days)']=pd.Series(time,index=production_df[frame].index)
	return production_df
