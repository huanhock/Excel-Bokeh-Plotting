#Module to be imported to master scripts for Bokeh Plotting Software
#Written By Kevin Zhai during Summer 2015 internship
#Modified: 8/15/15

import numpy as np

def stats(items, cl) :
	avg = np.mean(items)
	std = np.std(items)
	ucl = (cl * std) + avg
	lcl = avg - (cl * std)
	return {'avg':avg, 'std':std, 'ucl':ucl, 'lcl':lcl}

def liststats(items, cl) :
	if len(items) == 0:
		return {'avg':None, 'std':None, 'ucl':None, 'lcl':None}
	avg = []
	std = []
	for i in range(len(items[0])) : 
		temp = [x[i] for x in items] #list comprehenstion to obtain the ith element of each list and average them 
		avg.append(np.mean(temp))
		std.append(np.std(temp))
	ucl = [(x + (cl * y)) for x,y in zip(avg, std)] #list comprehension and zip function to add the std to the avg)
	lcl = [(x - (cl * y)) for x,y in zip(avg, std)]
	return {'avg':avg, 'std':std, 'ucl':ucl, 'lcl':lcl}

def normalize(items) :
	normalized = []
	for j in items:
		sixteen = j[8]
		temp = []
		for k in j:
			temp.append(round(k/sixteen,4))
		normalized.append(temp)
	return normalized

