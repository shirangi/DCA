# Author: Trent Bone
# Date: 2/16/2018

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

from DI_Downloads import *
from Data_Import_Functions import *
from PreProcessing import *

File=DI_Downloads()
nsheets=open_file(File)
df=DRI_to_Dataframe(File,nsheets)

script_dir=os.path.dirname(__file__)
workbook=xlwt.Workbook()
File_name="Karnes_GOR"
excel_sheet=workbook.add_sheet(File_name,True)
excel_sheet.write(0,1,'Well Name')
excel_sheet.write(0,2,'GOR_min')
excel_sheet.write(0,3,'GOR_max')
row=1


pbar=tqdm(total=nsheets)
for frame in xrange(0,nsheets):
	Oil=df[frame].loc[:,"Oil (STB/Month)"]
	Gas=df[frame].loc[:,"Gas (MSCF/Month)"]*1000
	GOR=np.array(Gas/Oil)
	GOR_min=np.nanmin(GOR)
	GOR_max=np.nanmax(GOR)
	excel_sheet.write(row,1,str('{:.0f}'.format(df[frame].loc[0,"API"])))
	excel_sheet.write(row,2,GOR_min)
	excel_sheet.write(row,3,GOR_max)
	row +=1
	pbar.update(1)
workbook.save('%s.xls' %(File_name))
