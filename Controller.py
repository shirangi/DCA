# Author: Trent Bone
# Date: 2/14/2018
from __future__ import
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

from Importer import *

# CONTROLLER
File=DI_Downloads()
nsheets=open_file(File)
df=DRI_to_Dataframe(File,nsheets)

pbar=tqdm(total=nsheets)
for frame in xrange(0,nsheets):
	T=df[frame].loc[:,"Delta_t (Days)"]
	enough_data=enough_data_check(r_max(df,frame))
	if enough_data==False:
		pbar.update(1)
		continue
	else:
		units,EUR_units,Q=Units(df,frame)
		sqrt_T,Np=Data_Prep(T,Q)
		T=sqrt_T**2
		params=Params_prep(r_max(df,frame),sqrt_T,Np)
		params=Regression(T,Np,params) # Returning list of variable answers
		[mcst,bBD,Npint,Npelf,telf]=params
		Decline=timeCum1(T,mcst,Npint,Npelf,bBD) # Using variables for Model
		# Percent Error
		########################################
		Error=Percent_Error(Np,T,mcst,Npint,Npelf,bBD)
		########################################
		# EUR Estimate
		########################################
		EUR=EURiabd(mcst, Npint, Npelf, bBD, 1)
		########################################
		Plotting(params,EUR_units,Decline,Error,EUR,sqrt_T,Np,df,frame)
		pbar.update(1)

