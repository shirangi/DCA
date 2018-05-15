# Author: Trent Bone
# Date: 2/14/2018
# Objective: Prepare data for fitting

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
############################

def r_max(df,frame):
	r_max=df[frame].shape[0] # Unpacks shape tuple for total rows
	return r_max

def enough_data_check(r_max):
	if r_max<12:
		return False
	else:
		return True

def Units(df,frame):
	if df[frame].loc[0,"Well Type"]=="OIL WELL":
		Q=df[frame].loc[:,"Oil (STB/Month)"]
		units="STB/Month"
		EUR_units="STB"
	else:
		Q=df[frame].loc[:,"Gas (MSCF/Month)"]
		units="MSCF/Month"
		EUR_units="MSCF"
	return units, EUR_units, Q

def data_to_array(Data):
	Data=np.asarray(Data)
	return Data

def tp_to_months(Data):
	Data=Data/30.4
	return Data

def time_to_tp(Data):
	# Turns time to cum time producing
	Data=np.cumsum(Data)
	return Data
def time_setup(Time):
	T=data_to_array(Time)
	T=time_to_tp(tp_to_months(T))
	sqrt_T=np.sqrt(T)
	return sqrt_T

def Production_Setup(Production):
	Q=data_to_array(Production)
	Np=np.cumsum(Q)
	return Np

def Data_Clean(T, Q):
	z=0
	while Q[z]==0:
		z=z+1
	Q=Q[z-1:]
	T=T[:len(T)-z+1]
	return T,Q

def Data_Prep(Time,Production):
	Time,Production=Data_Clean(Time,Production)
	sqrt_T=time_setup(Time)
	Np=Production_Setup(Production)
	return sqrt_T,Np




