#GTAT Grid Production Line Bokeh Plotting Software
#Written By Kevin Zhai during Summer 2015 internship
#Modified: 8/5/15
#Use with the Bokeh Plotting Software Guide
##############################################################################
#I. Importing

import statfunctions
import pandas as pd
import sys
from openpyxl import load_workbook
from bokeh.plotting import figure, output_file, show, gridplot, hplot
from bokeh.models import Callback, ColumnDataSource, GlyphRenderer, Circle, HoverTool
from bokeh.models.widgets import Panel, Tabs

##############################################################################

#II.Reading the Data 
wb = load_workbook('UpdatedGridThickness.xlsx') #read in excel file

cu12 = wb['Cu Tool 1 and 2']
meco = wb['MECO']

#CuTool1and2 Data
rows = sys.argv[1] #specifies the first argument
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

#Meco Tool Data
rows2 = sys.argv[2] #specifies the second argument 
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

####################################################################################
#III. Store data into Pandas Dataframe

cu12_data = {'id': pd.Series(cu12_id), 'tool': pd.Series(cu12_tool), 'weight': pd.Series(cu12_weight)} 
cu12_df = pd.DataFrame(cu12_data) #store data into panda dataframe

meco_data = {'id': pd.Series(meco_id), 'weight': pd.Series(meco_weight)}
meco_df = pd.DataFrame(meco_data)

rejected = cu1_rejected + cu2_rejected + meco_rejected #total number of rejected grids

#####################################################################################
#IV. Retrieve data From Pandas to be used in Bokeh
cl = 2 #control limit multiplier

cu1_ids = list(cu12_df[cu12_df['tool'] == 1].id)
cu1_weights = list(cu12_df[cu12_df['tool'] == 1].weight)
cu1_stats = statfunctions.stats(cu1_weights, cl)

cu2_ids = list(cu12_df[cu12_df['tool'] == 2].id)
cu2_weights = list(cu12_df[cu12_df['tool'] == 2].weight)
cu2_stats = statfunctions.stats(cu2_weights, cl)

m_ids = list(meco_df.id)
m_weights = list(meco_df.weight)
m_stats = statfunctions.stats(m_weights, cl)

c_thickness_stats = statfunctions.liststats(meco_c_thickness, cl)
l_thickness_stats = statfunctions.liststats(meco_l_thickness, cl)
r_thickness_stats = statfunctions.liststats(meco_r_thickness, cl)

######################################################################################
#V. Start Bokeh Plotting

first_items = first.split('/')
first_date = first_items[1] + '-' + first_items[0] + '-' + first_items[2]
last_items = last.split('/')
last_date = last_items[1] + '-' + last_items[0] + '-' + last_items[2]

if first_date == last_date:
	output_file(first_date + '.html', title=first_date)
else:
	output_file(first_date + ' to ' + last_date + '.html', title=first_date + ' to ' + last_date)


#Cu1 Plot
cu1_s = ColumnDataSource(data=dict(x=cu1_ids, y=cu1_weights))
p1 = figure(plot_width=780, plot_height=550, tools = "pan,box_select,box_zoom,xwheel_zoom,reset,save,resize", title="Cu Tool 1")
p1.line('x', 'y', source=cu1_s, line_width=2)
p1.circle('x', 'y', source=cu1_s, fill_color='white', size=8)

cu1_glyph = Circle(x='x',y='y', size=6, fill_color='white')
cu1_rend = GlyphRenderer(data_source=cu1_s, glyph=cu1_glyph)

cu1_hover = HoverTool(
        plot=p1,
        renderers=[cu1_rend],
        tooltips=[
            ("GridId, Weight", "@x, @y"),
        ]
    )

p1.tools.extend([cu1_hover])
p1.renderers.extend([cu1_rend])

cu1_avg = ColumnDataSource(data=dict(y=[cu1_stats['avg'], cu1_stats['avg']], cl=cl))
cu1_lower = ColumnDataSource(data=dict(y=[cu1_stats['lcl'], cu1_stats['lcl']]))
cu1_upper = ColumnDataSource(data=dict(y=[cu1_stats['ucl'], cu1_stats['ucl']]))
p1.line(x=[cu1_ids[0], cu1_ids[len(cu1_ids)-1]], y='y', source=cu1_avg, line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(cu1_stats['avg'], 3)))
p1.line(x=[cu1_ids[0], cu1_ids[len(cu1_ids)-1]], y='y', source=cu1_lower, line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(cu1_stats['std'], 3)) + ")")
p1.line(x=[cu1_ids[0], cu1_ids[len(cu1_ids)-1]], y='y', source=cu1_upper, line_width=1, line_color='red')

#javascript script to handle the selection interaction callback
cu1_s.callback = Callback(args=dict(cu1_avg=cu1_avg, cu1_lower=cu1_lower, cu1_upper=cu1_upper), code="""
        var inds = cb_obj.get('selected')['1d'].indices;
        var d = cb_obj.get('data');
        var items = [];

        if (inds.length == 0) { return; }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
        for (i = 0; i < inds.length; i++) {
            items.push(d['y'][inds[i]]);
        }

        var ym = average(items);
        var std = standardDeviation(items);

        cu1_avg.get('data')['y'] = [ym, ym]
        cu1_lower.get('data')['y'] = [ym - (cu1_avg.get('data')['cl'] * std), ym - (cu1_avg.get('data')['cl'] * std)]
        cu1_upper.get('data')['y'] = [(cu1_avg.get('data')['cl'] * std) + ym, (cu1_avg.get('data')['cl'] * std) + ym]

        cb_obj.trigger('change');
        cu1_avg.trigger('change');
        cu1_lower.trigger('change');
        cu1_upper.trigger('change');
    
    	//define the functions needed for standard dev and average
        function standardDeviation(values){
  			var avg = average(values);
  
  			var squareDiffs = values.map(function(value){
    			var diff = value - avg;
    			var sqrDiff = diff * diff;
    			return sqrDiff;
  			});
  
  			var avgSquareDiff = average(squareDiffs);

  			var stdDev = Math.sqrt(avgSquareDiff);
  			return stdDev;
		}

		function average(data){
  			var sum = data.reduce(function(sum, value){
   			 	return sum + value;
 			}, 0);

  			var avg = sum / data.length;
  			return avg;
		}
    """)

p1.line(cu1_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(cu1_ids))) #daily accepted
p1.line(cu1_ids, 1.2, line_width=1, legend='Rejected: ' + str(cu1_rejected)) #daily rejected
p1.line(cu1_ids, 1.2, line_width=2, line_color='green')
p1.line(cu1_ids, 1.4, line_width=2, line_color='green')
p1.background_fill = 'beige'
p1.xaxis.axis_label = 'Grid Id'
p1.yaxis.axis_label = 'Weights(g)'

#Cu2 Plot
cu2_s = ColumnDataSource(data=dict(x=cu2_ids, y=cu2_weights))
p2 = figure(plot_width=780, plot_height=550, tools = "pan,box_select,box_zoom,xwheel_zoom,reset,save,resize", title="Cu Tool 2")
p2.line('x', 'y', source=cu2_s, line_width=2)
p2.circle('x', 'y', source=cu2_s, fill_color='white', size=8)

cu2_glyph = Circle(x='x',y='y', size=6, fill_color='white')
cu2_rend = GlyphRenderer(data_source=cu2_s, glyph=cu2_glyph)

cu2_hover = HoverTool(
        plot=p2,
        renderers=[cu2_rend],
        tooltips=[
            ("GridId, Weight", "@x, @y"),
        ]
    )
p2.tools.extend([cu2_hover])
p2.renderers.extend([cu2_rend])

cu2_avg = ColumnDataSource(data=dict(y=[cu2_stats['avg'], cu2_stats['avg']], cl=cl))
cu2_lower = ColumnDataSource(data=dict(y=[cu2_stats['lcl'], cu2_stats['lcl']]))
cu2_upper = ColumnDataSource(data=dict(y=[cu2_stats['ucl'], cu2_stats['ucl']]))
if len(cu2_ids) != 0:
  p2.line(x=[cu2_ids[0], cu2_ids[len(cu2_ids)-1]], y='y', source=cu2_avg, line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(cu2_stats['avg'], 3)))
  p2.line(x=[cu2_ids[0], cu2_ids[len(cu2_ids)-1]], y='y', source=cu2_lower, line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(cu2_stats['std'], 3)) + ")")
  p2.line(x=[cu2_ids[0], cu2_ids[len(cu2_ids)-1]], y='y', source=cu2_upper, line_width=1, line_color='red')

#javascript script to handle the selection interaction callback
cu2_s.callback = Callback(args=dict(cu2_avg=cu2_avg, cu2_lower=cu2_lower, cu2_upper=cu2_upper), code="""
        var inds = cb_obj.get('selected')['1d'].indices;
        var d = cb_obj.get('data');
        var items = [];

        if (inds.length == 0) { return; }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
        for (i = 0; i < inds.length; i++) {
            items.push(d['y'][inds[i]]);
        }

        var ym = average(items);
        var std = standardDeviation(items);

        cu2_avg.get('data')['y'] = [ym, ym]
        cu2_lower.get('data')['y'] = [ym - (cu2_avg.get('data')['cl'] * std), ym - (cu2_avg.get('data')['cl'] * std)]
        cu2_upper.get('data')['y'] = [(cu2_avg.get('data')['cl'] * std) + ym, (cu2_avg.get('data')['cl'] * std) + ym]

        cb_obj.trigger('change');
        cu2_avg.trigger('change');
        cu2_lower.trigger('change');
        cu2_upper.trigger('change');
    
      //define the functions needed for standard dev and average
        function standardDeviation(values){
        var avg = average(values);
  
        var squareDiffs = values.map(function(value){
          var diff = value - avg;
          var sqrDiff = diff * diff;
          return sqrDiff;
        });
  
        var avgSquareDiff = average(squareDiffs);

        var stdDev = Math.sqrt(avgSquareDiff);
        return stdDev;
    }

    function average(data){
        var sum = data.reduce(function(sum, value){
          return sum + value;
      }, 0);

        var avg = sum / data.length;
        return avg;
    }
    """)

p2.line(cu2_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(cu2_ids))) #daily accepted
p2.line(cu2_ids, 1.2, line_width=1, legend='Rejected: ' + str(cu2_rejected)) #daily rejected
p2.line(cu2_ids, 1.2, line_width=2, line_color='green')
p2.line(cu2_ids, 1.4, line_width=2, line_color='green')
p2.background_fill = 'beige'
p2.xaxis.axis_label = 'Grid Id'
p2.yaxis.axis_label = 'Weights(g)'


#Meco plot
m_s = ColumnDataSource(data=dict(x=m_ids, y=m_weights))
p3 = figure(plot_width=780, plot_height=550, tools = "pan,box_select,box_zoom,xwheel_zoom,reset,save,resize", title="Meco Tool")
p3.line('x', 'y', source=m_s, line_width=2)
p3.circle('x', 'y', source=m_s, fill_color='white', size=8)

m_glyph = Circle(x='x',y='y', size=6, fill_color='white')
m_rend = GlyphRenderer(data_source=m_s, glyph=m_glyph)

m_hover = HoverTool(
        plot=p3,
        renderers=[m_rend],
        tooltips=[
            ("GridId, Weight", "@x, @y"),
        ]
    )
p3.tools.extend([m_hover])
p3.renderers.extend([m_rend])

m_avg = ColumnDataSource(data=dict(y=[m_stats['avg'], m_stats['avg']], cl=cl))
m_lower = ColumnDataSource(data=dict(y=[m_stats['lcl'], m_stats['lcl']]))
m_upper = ColumnDataSource(data=dict(y=[m_stats['ucl'], m_stats['ucl']]))
p3.line(x=[m_ids[0], m_ids[len(m_ids)-1]], y='y', source=m_avg, line_width=2, line_color='red', line_dash=[4,4], legend='Mean = ' + str(round(m_stats['avg'], 3)))
p3.line(x=[m_ids[0], m_ids[len(m_ids)-1]], y='y', source=m_lower, line_width=1, line_color='red', legend='2*Std (Std = ' + str(round(m_stats['std'], 3)) + ")")
p3.line(x=[m_ids[0], m_ids[len(m_ids)-1]], y='y', source=m_upper, line_width=1, line_color='red')

#javascript script to handle the selection interaction callback
m_s.callback = Callback(args=dict(m_avg=m_avg, m_lower=m_lower, m_upper=m_upper), code="""
        var inds = cb_obj.get('selected')['1d'].indices;
        var d = cb_obj.get('data');
        var items = [];

        if (inds.length == 0) { return; }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
        for (i = 0; i < inds.length; i++) {
            items.push(d['y'][inds[i]]);
        }

        var ym = average(items);
        var std = standardDeviation(items);

        m_avg.get('data')['y'] = [ym, ym]
        m_lower.get('data')['y'] = [ym - (m_avg.get('data')['cl'] * std), ym - (m_avg.get('data')['cl'] * std)]
        m_upper.get('data')['y'] = [(m_avg.get('data')['cl'] * std) + ym, (m_avg.get('data')['cl'] * std) + ym]

        cb_obj.trigger('change');
        m_avg.trigger('change');
        m_lower.trigger('change');
        m_upper.trigger('change');
    
    	//define the functions needed for calculating standard dev and average
        function standardDeviation(values){
  			var avg = average(values);
  
  			var squareDiffs = values.map(function(value){
    			var diff = value - avg;
    			var sqrDiff = diff * diff;
    			return sqrDiff;
  			});
  
  			var avgSquareDiff = average(squareDiffs);

  			var stdDev = Math.sqrt(avgSquareDiff);
  			return stdDev;
		}

		function average(data){
  			var sum = data.reduce(function(sum, value){
   			 	return sum + value;
 			}, 0);

  			var avg = sum / data.length;
  			return avg;
		}
    """)

p3.line(m_ids, 1.4, line_width=1, legend='Accepted: ' + str(len(m_ids))) #daily accepted
p3.line(m_ids, 1.2, line_width=1, legend='Rejected: ' + str(meco_rejected)) #daily rejected
p3.line(m_ids, 1.2, line_width=2, line_color='green')
p3.line(m_ids, 1.4, line_width=2, line_color='green')
p3.background_fill = 'beige'
p3.xaxis.axis_label = 'Grid Id'
p3.yaxis.axis_label = 'Weights(g)'


#Meco GridThicknesses
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

#Normalize Center Thicknesses
normalized_thickness = statfunctions.normalize(meco_c_thickness)
n_thickness_stats = statfunctions.liststats(normalized_thickness, cl)

p4b = figure(plot_width=700, plot_height=400, title="Center Thickness Control (Normalized)")
p4b.line(center_measurements, n_thickness_stats['avg'], line_width=2, legend='Center')
p4b.circle(center_measurements, n_thickness_stats['avg'], fill_color='white', size=8)
#STD lines
p4b.line(center_measurements, n_thickness_stats['ucl'], line_width=1, line_color='red', legend='2 Std Lines')
p4b.line(center_measurements, n_thickness_stats['lcl'], line_width=1, line_color='red')
p4b.xaxis.axis_label = 'Measurement Number'
p4b.yaxis.axis_label = 'Thickness (mm)'

#create Panel layout
tab1 = Panel(child=hplot(p1, p2), title="Cu Tool Weights") #first tab of the dashboard
tab2 = Panel(child=p3, title="MECO Weights") #second tab of the dashboard
tab3 = Panel(child=hplot(p4a, p4b), title="MECO Thicknesses")

tabs = Tabs(tabs=[tab1, tab2, tab3])
show(tabs)