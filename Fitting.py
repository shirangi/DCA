# Author: Trent Bone
# Date: 2/14/2018
# Objective: Fitting data and returning all variable values.

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

def Params_prep(r_max,sqrt_T,Np):
	guess=int(r_max * 0.9)
	mcst=slope(sqrt_T[0:guess],Np[0:guess])[0]
	Npint=slope(sqrt_T[0:guess],Np[0:guess])[1]
	Npelf_max=np.nanmax(Np)
	Npelf_average=np.median(Np)
	params=Parameters()
	params.add('mcst',value=mcst, min=0, vary=True)
	params.add('Npint',value=-1, max=0, vary=True)
	params.add('Npelf', value=Npelf_average, min=0, max=Npelf_max, vary=True)
	params.add('bBD',value=0.1,min=0, max=1,vary=True)
	return params

def residual(params,T,Np):
	parvals=params.valuesdict()
	mcst=parvals['mcst']
	Npint=parvals['Npint']
	Npelf=parvals['Npelf']
	bBD=parvals['bBD']
	model=timeCum1(T,mcst,Npint,Npelf,bBD)
	model=np.nan_to_num(model)
	residual=abs(Np-model)
	for i in xrange(0,len(Np)):
		if Np[i]==0:
			residual[i]=0
	return residual.view(np.float)

def Regression(T,Np,params):
	# MODELING THE ARPS W/ DATA
	[mcst,bBD,Npint,Npelf]=params
	result=minimize(residual,params,method='leastsq',args=(T,Np))
	# Turning Best_Fit Dictionary into variables.
	locals().update(result.params)
	parvals=result.params.valuesdict()
	mcst=parvals['mcst']
	Npint=parvals['Npint']
	Npelf=parvals['Npelf']
	bBD=parvals['bBD']
	telf=(Npelf-Npint)/mcst
	params=[mcst,bBD,Npint,Npelf,telf]
	return params

def slope(x,y):
	slope=stats.linregress(x,y)
	return slope

def EURiabd(mcsrt, Npint, Npelf, bBD, qecl):
	# BD, boundary dominated
	# deterministic; written in oil terms but also for gas replacing Npelf with Gpelf
	# maximum life is 40 years
	# units monthly
	# utilize slope and intercept from Cumulative vs square root of time
	telf = ((Npelf - Npint) / mcsrt) ** 2
	qelf = mcsrt / (2 * (telf ** 0.5))
	# determine D0 and Delf
	if Npint == 0: # infinite conductivity
		Delf = 1 / (2 * telf)
	else: # finite conductivity
		D0 = 0.5 * (mcsrt / Npint) ** 2
		Delf = 1 / (1 / D0 + 2 * telf)
	NpBD = hyperbolicCumRate(qelf, Delf, bBD, qecl)
	# check for life > 40 years
	maxBDlife = 40 * 12 - telf # total time (in months) during boundary dominated flow
	NpBDL = hyperbolicCum(qelf, Delf, bBD, maxBDlife) # cumulative production during max life
	if NpBDL < NpBD:
		NpBD = NpBDL
	EURiabd = Npelf + NpBD
	EURiabd=format(EURiabd,'.0f')
	return EURiabd

def timeCum1(time, mcst, Npint, Npelf, bBD):
	#for cumulative vs square of time
	# Npelf for primary phase, oil or gas
	n=time.size
	timeCum1=np.zeros(n)
	i=0
	mrrc = 2 / (mcst**2)
	telf = ((Npelf - Npint) / mcst) ** 2 # back calculation
	while (i < n):
		if (time[i] < telf):
			timeCum1[i] = Npint + mcst * (time[i] ** 0.5)
		else:
			qelf = 0.5 * mcst / (telf ** 0.5) # first deriviative
			
			if (Npint == 0): # infinite conductivity
				Delf = 1 / (2 * telf) # dividing nominal equation numerator and denominator by Di
			else:
			# finite conductivity
				Delf = mcst / qelf / 4 * (telf ** (-3 / 2)) # second deriviative
				timeCum1[i] = hyperbolicCum(qelf, Delf, bBD, time[i] - telf) + Npelf
		i=i+1
	return timeCum1

def hyperbolicCumRate(qi, Di, b, q):
# cum as a function of rate
	if b == 1:
		hyperbolicCumRate = harmonicCumRate(qi, Di, q)
		return
	else:
		hyperbolicCumRate = qi ** b / (Di * (1 - b)) * (qi ** (1 - b) - q ** (1 - b))
		return hyperbolicCumRate


def hyperbolicCum(qi, di, b, t):
	# works for harmonic
	if (b == 1):
		hyperbolicCum = harmonicCum(qi, di, t)
	if (b==0):
		q=exponentialRate(qi,di,t)
		hyperbolicCum = qi ** b / (di * (1 - b)) * (qi ** (1 - b) - q ** (1 - b))
	else:
		q = hyperbolicRate(qi, di, b, t)
		hyperbolicCum = qi ** b / (di * (1 - b)) * (qi ** (1 - b) - q ** (1 - b))
	return hyperbolicCum

def harmonicCum(qi, di, t):
	q = qi / (1 + di * t)
	harmonicCum = qi / di * math.log(qi / q)
	return harmonicCum

def exponentialRate(qi, di, t):
	exponentialRate = qi * math.exp(-di * t)
	return exponentialRate

def hyperbolicRate(qi, di, b,t):
	if (b == 0): 
		hyperbolicRate = exponentialRate(qi, di, t)
	else:
		hyperbolicRate = qi / (1 + b * di * t) ** (1 / b)
	return hyperbolicRate

def hyperbolicLife(qi, di, b, qecl):
	if b == 0:
		hyperbolicLife = exponetialLife(qi, di, qecl)
		return hyperbolicLife
	hyperbolicLife = ((qi / qecl) ** b - 1) / (b * di)
	return hyperbolicLife

def exponetialLife(qi, D, qecl):
	exponetialLife = math.log(qi / qecl) / D
	return exponetialLife

def Percent_Error(Np,T,mcst,Npint,Npelf,bBD):
	model=timeCum1(T,mcst,Npint,Npelf,bBD)
	Error=np.average(abs((Np-model)/model))*100
	Error=format(Error,'.2f')
	return Error
