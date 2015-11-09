#GTAT Grid Production Line Bokeh Plotting Software
#Written By Kevin Zhai during Summer 2015 internship
#Modified: 8/13/15
#Use with the Bokeh Plotting Software Guide
##############################################################################

import statfunctions
from openpyxl import load_workbook
from bokeh.plotting import figure, output_file, show, gridplot, hplot
from bokeh.charts import Bar
import sys
from bokeh.models.widgets import Panel, Tabs
import pandas as pd

wb = load_workbook('UpdatedGridThickness.xlsx')

cu12 = wb['Cu Tool 1 and 2']
meco = wb['MECO']

#CuTool1and2 Data
rows = sys.argv[1]
dates = rows.split(':') 
first = cu12['A' + dates[0]].value
last = cu12['A' + dates[1]].value

current = int(dates[0])
#Counter for number of rejected parts (X's, Scraps, etc.) for each tool
cu1_rejected = 0
cu2_rejected = 0
meco_rejected = 0

cu12_id = []
cu12_tool = []
cu12_weight = []

#Retrieve wanted data from spreadsheet
while current <= int(dates[1]):
	if type(cu12['AG' + str(current)].value) is float:
		cu12_id.append(cu12['M' + str(current)].value)
		cu12_tool.append(int(cu12['L' + str(current)].value))
		cu12_weight.append(cu12['AG' + str(current)].value)
		current += 1
	else:
		if cu12['L' + str(current)].value == '01':
			cu1_rejected += 1
			current += 1
		else:
			cu2_rejected += 1
			current += 1

cu12_data = {'id': pd.Series(cu12_id), 'tool': pd.Series(cu12_tool), 'weight': pd.Series(cu12_weight)} 
cu12_df = pd.DataFrame(cu12_data) #store data into panda dataframe

##################################################################################
#Meco Tool Data
rows2 = sys.argv[2]
dates2 = rows2.split(':')

current2 = int(dates2[0])

meco_id = []
meco_weight = []
meco_c_thickness = [] 
meco_l_thickness = []
meco_r_thickness = [] 

#Retrieve wanted data from spreadsheet
while current2 <= int(dates2[1]):
	if type(meco['AF' + str(current2)].value) is float: #only add the grid if it is not broken
		meco_id.append(meco['L' + str(current2)].value)
		meco_weight.append(meco['AF' + str(current2)].value)
		if type(meco['P' + str(current2)].value) is float: #only add the thickness if it is a valid measurement
			meco_c_thickness.append([meco['P' + str(current2)].value, meco['S' + str(current2)].value, meco['U' + str(current2)].value, meco['V' + str(current2)].value, meco['X' + str(current2)].value, meco['Z' + str(current2)].value, meco['AA' + str(current2)].value, meco['AB' + str(current2)].value, meco['AD' + str(current2)].value])
			meco_l_thickness.append([meco['O' + str(current2)].value, meco['R' + str(current2)].value, meco['W' + str(current2)].value, meco['AC' + str(current2)].value])
			meco_r_thickness.append([meco['Q' + str(current2)].value, meco['T' + str(current2)].value, meco['Y' + str(current2)].value, meco['AE' + str(current2)].value])
		current2 += 1
	else:
		meco_rejected += 1
		current2 += 1

meco_data = {'id': pd.Series(meco_id), 'weight': pd.Series(meco_weight)}
meco_df = pd.DataFrame(meco_data)

rejected = cu1_rejected + cu2_rejected + meco_rejected #total number of rejected grids
#####################################################################################
#Format Pandas df into list to be used in Bokeh

cu1_ids = list(cu12_df[cu12_df['tool'] == 1].id)
cu1_weights = list(cu12_df[cu12_df['tool'] == 1].weight)
cu1_stats = statfunctions.stats(cu1_weights)

cu2_ids = list(cu12_df[cu12_df['tool'] == 2].id)
cu2_weights = list(cu12_df[cu12_df['tool'] == 2].weight)
cu2_stats = statfunctions.stats(cu2_weights)

m_ids = list(meco_df.id)
m_weights = list(meco_df.weight)
m_stats = statfunctions.stats(m_weights)

c_thickness_stats = statfunctions.liststats(meco_c_thickness)
l_thickness_stats = statfunctions.liststats(meco_l_thickness)
r_thickness_stats = statfunctions.liststats(meco_r_thickness)

######################################################################################
#Start Bokeh Plotting

first_items = first.split('/')
first_date = first_items[1] + '-' + first_items[0] + '-' + first_items[2]
last_items = last.split('/')
last_date = last_items[1] + '-' + last_items[0] + '-' + last_items[2]

if first_date == last_date:
	output_file(first_date + '.html', title=first_date)
else:
	output_file(first_date + ' to ' + last_date + '.html', title=first_date + ' to ' + last_date)

#Cu1 Plot
p1 = figure(plot_width=700, plot_height=400, title="Cu Tool 1")
p1.line(cu1_ids, cu1_weights, line_width=2)
p1.circle(cu1_ids, cu1_weights, fill_color='white', size=8)
p1.line(cu1_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(cu1_ids))) #daily accepted
p1.line(cu1_ids, 1.2, line_width=1, legend='Rejected: ' + str(cu1_rejected)) #daily rejected
p1.line(cu1_ids, cu1_stats['avg'], line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(cu1_stats['avg'], 3)))
p1.line(cu1_ids, cu1_stats['lcl'], line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(cu1_stats['std'], 3)) + ")")
p1.line(cu1_ids, cu1_stats['ucl'], line_width=1, line_color='red')
p1.line(cu1_ids, 1.2, line_width=2, line_color='green')
p1.line(cu1_ids, 1.4, line_width=2, line_color='green')
p1.background_fill = 'beige'
p1.xaxis.axis_label = 'Grid Id'
p1.yaxis.axis_label = 'Weights(g)'

#Cu2 plot
p2 = figure(plot_width=700, plot_height=400, title="Cu Tool 2")
p2.line(cu2_ids, cu2_weights, line_width=2)
p2.circle(cu2_ids, cu2_weights, fill_color='white', size=8)
p2.line(cu2_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(cu2_ids))) #daily accepted
p2.line(cu2_ids, 1.2, line_width=1, legend='Rejected: ' + str(cu2_rejected)) #daily rejected
p2.line(cu2_ids, cu2_stats['avg'], line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(cu2_stats['avg'], 3)))
p2.line(cu2_ids, cu2_stats['lcl'], line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(cu2_stats['std'], 3)) + ")")
p2.line(cu2_ids, cu2_stats['ucl'], line_width=1, line_color='red')
p2.line(cu2_ids, 1.2, line_width=2, line_color='green')
p2.line(cu2_ids, 1.4, line_width=2, line_color='green')
p2.background_fill = 'beige'
p2.xaxis.axis_label = 'Grid Id'
p2.yaxis.axis_label = 'Weights(g)'

#Meco plot
p3 = figure(plot_width=700, plot_height=400, title="Meco Tool")
p3.line(m_ids, m_weights, line_width=2)
p3.circle(m_ids, m_weights, fill_color='white', size=8)
p3.line(m_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(m_ids))) #daily accepted
p3.line(m_ids, 1.2, line_width=1, legend='Rejected: ' + str(meco_rejected)) #daily rejected
p3.line(m_ids, m_stats['avg'], line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(m_stats['avg'], 3)))
p3.line(m_ids, m_stats['lcl'], line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(m_stats['std'], 3)) + ")")
p3.line(m_ids, m_stats['ucl'], line_width=1, line_color='red')
p3.line(m_ids, 1.2, line_width=2, line_color='green')
p3.line(m_ids, 1.4, line_width=2, line_color='green')
p3.background_fill = 'beige'
p3.xaxis.axis_label = 'Grid Id'
p3.yaxis.axis_label = 'Weights(g)'


#Meco Center Thicknesses
p4a = figure(plot_width=700, plot_height=400, title="Meco Thickness Comparison")
center_measurements = [2,5,7,8,10,12,13,14,16]
l_measurements = [1,4,9,15]
r_measurements = [3,6,11,17]
#center thickness
p4a.line(center_measurements, c_thickness_stats['avg'], line_width=2, legend='Center')
p4a.circle(center_measurements, c_thickness_stats['avg'], fill_color='white', size=8)
#left thickness
p4a.line(l_measurements, l_thickness_stats['avg'], line_width=2, legend='Left', color='orange')
p4a.circle(l_measurements, l_thickness_stats['avg'], fill_color='white', size=8)
#right thickness
p4a.line(r_measurements, r_thickness_stats['avg'], line_width=2, legend='Right', color='purple')
p4a.circle(r_measurements, r_thickness_stats['avg'], fill_color='white', size=8)

p4a.xaxis.axis_label = 'Measurement Number'
p4a.yaxis.axis_label = 'Thickness (mm)'
tab1 = Panel(child=p4a, title="Thickness Comparisons")

p4b = figure(plot_width=700, plot_height=400, title="Center Thickness Control")
p4b.line(center_measurements, c_thickness_stats['avg'], line_width=2, legend='Center')
p4b.circle(center_measurements, c_thickness_stats['avg'], fill_color='white', size=8)
#STD lines
p4b.line(center_measurements, c_thickness_stats['ucl'], line_width=1, line_color='red', legend='2 Std Lines')
p4b.line(center_measurements, c_thickness_stats['lcl'], line_width=1, line_color='red')
p4b.xaxis.axis_label = 'Measurement Number'
p4b.yaxis.axis_label = 'Thickness (mm)'
tab2 = Panel(child=p4b, title="Thickness Control")

p4_tabs = Tabs(tabs=[tab1, tab2])


p = gridplot([[p1, p2], [p3, p4a]])
show(p)
