
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

from Fitting import *

def Plotting(params,EUR_units,Decline,Error,EUR,sqrt_T,Np,df,frame):
	[mcst,bBD,Npint,Npelf,telf]=params
	########## GETTING PATH AND FILE########################
	script_dir=os.path.dirname(__file__)
	#######################################################	

	## PLOTTING
	############################################
	plt.xlabel("Sqrt[Time] (Months)", fontsize=15)
	plt.ylabel("Cumulative Production %s" %(EUR_units), fontsize=15)
	Decline=plt.plot(sqrt_T,Decline,color="red", linewidth=2, alpha=1, label="Model")
	plt.ticklabel_format(fontsize=25)
	plt.plot(sqrt_T,Np,'bo', label="Data", markerfacecolor="None", markeredgecolor='blue',markeredgewidth=0.5)
	plt.plot(sqrt_T,Np,'ro',alpha=0,label="mcst=%s" %(format(mcst,'.0f'))) # Plots invisible point that I use to just have the proper info in the legend
	plt.plot(sqrt_T,Np,'ro',alpha=0,label="Npint=%s" %(format(Npint,'.0f'))) # Plots invisible point that I use to just have the proper info in the legend
	plt.plot(telf,Npelf,'y^',alpha=1,label="Npelf=%s" %(format(Npelf,'.0f'))) # Plots invisible point that I use to just have the proper info in the legend
	plt.plot(sqrt_T,Np,'ro',alpha=0,label="bBD=%s" %(format(bBD,'.2f'))) # Plots invisible point that I use to just have the proper info in the legend
	plt.plot(sqrt_T,Np,'ro',alpha=0,label="EUR=%s %s" %(EUR,EUR_units)) # Plots invisible point that I use to just have the proper info in the legend
	plt.plot(sqrt_T,Np,'ro',alpha=0,label="Percent Error=%s %s" %(Error,'%')) # Plots invisible point that I use to just have the proper info in the legend
	plt.legend(loc='upper left')
	name=str('{:.0f}'.format(df[frame].loc[0,"API"]))
	plt.savefig('%s\%s.png' %(script_dir,name), dpi=None, facecolor='w', edgecolor='w',orientation='portrait', papertype=None, format=None,transparent=False, bbox_inches=None, pad_inches=0.1,frameon=None)
	plt.gcf().clear()
	#################################
