# DCA
Decline Curve Analysis Tools

AUTHOR: Trent Bone
OBJECTIVES: 
	1). Take production data from drillinginfo and turn it into usable data for python
	2). Classify the flow regime of each well using machine learning.
	3). Do decline curve analysis on each well based on what flow regime it is in
		-> IABD (Hyperbolic Arps)
		-> IA	(Monte Carlo)
		-> LPL	(Thesis or Monte Carlo)
		-> LPLBD(Hyperbolic Arps)
		-> Other(Ignore or maybe use Arps best fit)
	4). Write all data into a format that is easily importable to various softwares to do calculations or contours.

MODULES:
	1). DI_Downloads
		-> Takes DRI files and puts them into an excel file and deletes copies
	2). DCA_Database
		-> Test for pandas dataframe
	3). DCA_Multi
		-> Arps Hyperbolic fitting for IABD wells (fitting needs improving).
	4). Arps_Hyperbolic
		->Arps singular best fit
