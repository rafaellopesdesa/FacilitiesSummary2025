import numpy as np
import pandas as pd
from scipy.optimize import minimize
import gspread
import time
import gspread_formatting
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.lines import Line2D

def readInputs(inputs_sheet):
  inputs = {}
  inputs['computing'] = {}
  inputs['storage'] = {}
  inputs['resources'] = {}

  resources = inputs_sheet.col_values(1)
  keys = inputs_sheet.col_values(2)

  row = 0
  for resource,key in zip(resources,keys):
    row = row + 1
    if 'target' in key.lower():
      inputs[resource.lower()]['target'] = [float(val) for val in inputs_sheet.row_values(row)[2:]]
    elif 'cost' in key.lower():
      inputs[resource.lower()]['cost'] = [float(val) for val in inputs_sheet.row_values(row)[2:]]
    if 'years' in key.lower():
      inputs[resource.lower()]['years'] = [float(val) for val in inputs_sheet.row_values(row)[2:]]
    elif 'lifetime' in key.lower():
      inputs[resource.lower()]['lifetime'] = [float(val) for val in inputs_sheet.row_values(row)[2:]]
  return inputs

def readSite(site_sheet):
  site = {}
  site['computing'] = {}
  site['storage'] = {}
  site['cost'] = {}
  site['resources'] = {}

  resources = site_sheet.col_values(1)
  keys = site_sheet.col_values(2)

  row = 0
  for resource,key in zip(resources,keys):
    row = row + 1
    if 'initial' in key.lower():
      site[resource.lower()]['initial'] = [float(val) for val in site_sheet.row_values(row)[2:]]
    if 'retirement' in key.lower():
      site[resource.lower()]['retirement'] = [float(val) for val in site_sheet.row_values(row)[2:]]
    if 'years' in key.lower():
      site[resource.lower()]['years'] = [float(val) for val in site_sheet.row_values(row)[2:]]
    if 'size' in key.lower():
      site[resource.lower()]['size'] = [float(val) for val in site_sheet.row_values(row)[2:]]
    if 'non-equipment' in key.lower():
      site[resource.lower()]['nonequipment'] = [float(val) for val in site_sheet.row_values(row)[2:]]
  return site

def readScenario(scenario_sheet):
  scenario = {}
  scenario['cost'] = {}
  scenario['resources'] = {}
  resources = scenario_sheet.col_values(1)
  keys = scenario_sheet.col_values(2)

  row = 0
  for resource,key in zip(resources,keys):
    row = row + 1

    if 'years' in key.lower():
      scenario[resource.lower()]['years'] = [float(val) for val in scenario_sheet.row_values(row)[2:]]
    if 'budget' in key.lower():
      scenario[resource.lower()]['budget'] = [float(val) for val in scenario_sheet.row_values(row)[2:]]
  return scenario

def cost_function(fraction, inputs, site, scenario, debug):

  loss = 0.0
  debug.clear()
  
  storage = site['storage']['initial'][0]
  computing = site['computing']['initial'][0]
  size = site['resources']['size'][0]

  storage_oldjunk = 0
  computing_oldjunk = 0

  for year, budget, storage_fraction in \
  zip(scenario['resources']['years'], scenario['cost']['budget'], fraction):

    inputs_index = inputs['resources']['years'].index(year)
    site_index = site['resources']['years'].index(year)

    budget = max(budget*size - site['cost']['nonequipment'][site_index],0)

    storage_cost = inputs['storage']['cost'][inputs_index]
    storage_retirement = site['storage']['retirement'][site_index] + storage_oldjunk
    storage_target = inputs['storage']['target'][inputs_index]*size

    computing_cost = inputs['computing']['cost'][inputs_index]
    computing_retirement = site['computing']['retirement'][site_index] + computing_oldjunk
    computing_target = inputs['computing']['target'][inputs_index]*size
    
    storage = max(storage - storage_retirement,0.0)
    storage = storage + storage_fraction*budget/storage_cost
    computing = max(computing - computing_retirement,0.0) 
    computing = computing + (1-storage_fraction)*budget/computing_cost

    storage_oldjunk = max(min(storage_retirement,storage_target - storage), 0)
    computing_oldjunk = max(min(computing_retirement, computing_target - computing), 0)
    
    storage = storage + storage_oldjunk
    computing = computing + computing_oldjunk
    
    loss += ((storage_target - storage)/storage_target)**2
    loss += ((computing_target - computing)/computing_target)**2
    loss += 0.5*(storage_oldjunk/storage_target)**2
    loss += 0.5*(computing_oldjunk/computing_target)**2

    debug.append((storage_fraction*budget,storage_target, storage, storage_oldjunk, \
    (1-storage_fraction)*budget, computing_target, computing, computing_oldjunk))
  
  return loss

def writeReport(output_worksheet, site, scenario, scenario_name, debug, initial_row):

  title_reference = gspread.utils.rowcol_to_a1(initial_row, 1)
  bad_color = gspread_formatting.cellFormat(backgroundColor=gspread_formatting.Color(0.576, 0.80, 0.918))
  good_color = gspread_formatting.cellFormat(backgroundColor=gspread_formatting.Color(0.851, 0.918, 0.828))

  output_worksheet.update_cell(initial_row, 1, scenario_name)
  output_worksheet.format(title_reference, {'textFormat': {'bold': True}})
  output_worksheet.update_cell(initial_row, 3, 'Initial')
  for i,year in enumerate(scenario['resources']['years']):
    output_worksheet.update_cell(initial_row, i+4, year)
  output_worksheet.update_cell(initial_row+1, 1, 'Storage')
  output_worksheet.update_cell(initial_row+1, 2, 'Total (TB)')
  output_worksheet.update_cell(initial_row+2, 2, 'Target (TB)')
  output_worksheet.update_cell(initial_row+3, 2, 'Budget ($)')
  output_worksheet.update_cell(initial_row+4, 2, 'Junk (TB)')
  output_worksheet.update_cell(initial_row+1, 3, site['storage']['initial'][0])
  output_worksheet.update_cell(initial_row+5, 1, 'Computing')
  output_worksheet.update_cell(initial_row+6, 2, 'Total (HS23)')
  output_worksheet.update_cell(initial_row+7, 2, 'Target (HS23)')
  output_worksheet.update_cell(initial_row+8, 2, 'Budget ($)')
  output_worksheet.update_cell(initial_row+9, 2, 'Junk (HS23)')
  output_worksheet.update_cell(initial_row+6, 3, site['computing']['initial'][0])
  for i, data in enumerate(debug):
    output_worksheet.update_cell(initial_row+1, i+4, data[2])

    cell_reference = gspread.utils.rowcol_to_a1(initial_row+1, i+4)
    if data[2] < data[1]:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, bad_color)
    else:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, good_color)

    output_worksheet.update_cell(initial_row+2, i+4, data[1])
    output_worksheet.update_cell(initial_row+3, i+4, data[0])
    output_worksheet.update_cell(initial_row+4, i+4, data[3])
    
    cell_reference = gspread.utils.rowcol_to_a1(initial_row+4, i+4)
    if data[3] > 0.1*data[1]:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, bad_color)
    else:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, good_color)
    
    output_worksheet.update_cell(initial_row+6, i+4, data[6])
    
    cell_reference = gspread.utils.rowcol_to_a1(initial_row+6, i+4)
    if data[6] < data[5]:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, bad_color)
    else:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, good_color)
    
    output_worksheet.update_cell(initial_row+7, i+4, data[5])
    output_worksheet.update_cell(initial_row+8, i+4, data[4])
    output_worksheet.update_cell(initial_row+9, i+4, data[7])

    cell_reference = gspread.utils.rowcol_to_a1(initial_row+9, i+4)
    if data[7] > 0.1*data[5]:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, bad_color)
    else:
      gspread_formatting.format_cell_range(output_worksheet, cell_reference, good_color)
    

def optimize(spreadsheet, sites, outputs):
  inputs = readInputs(spreadsheet.worksheet('Inputs'))
  retval = []

  for site, output in zip(sites, outputs):

    sheets = spreadsheet.worksheets()
    site = readSite(spreadsheet.worksheet(site))
    output = spreadsheet.worksheet(output)
    output.clear()

    line = 1
    scenarios_debug = []
    for scenario_sheet in sheets:

      if 'Scenario' in scenario_sheet.title:
        debug = []
        scenario = readScenario(scenario_sheet)
        years = len(scenario['resources']['years'])
        fractions = years*[0.5]
        bounds = years*[(0, 1)]
        minimize(cost_function, fractions, bounds=bounds, args=(inputs, site, scenario, debug))
        writeReport(output, site, scenario, scenario_sheet.title, debug, line)
        scenarios_debug.append([scenario['resources']['years'],debug])
        line = line + 11
    retval.append(scenarios_debug)
    time.sleep(5)
  return retval

def multicolor_ylabel(ax,list_of_strings,list_of_colors,axis='x',anchorpad=0,**kw):
    """this function creates axes labels with multiple colors
    ax specifies the axes object where the labels should be drawn
    list_of_strings is a list of all of the text items
    list_if_colors is a corresponding list of colors for the strings
    axis='x', 'y', or 'both' and specifies which label(s) should be drawn"""
    from matplotlib.offsetbox import AnchoredOffsetbox, TextArea, HPacker, VPacker

    # x-axis label
    if axis=='x' or axis=='both':
        boxes = [TextArea(text, textprops=dict(color=color, ha='left',va='bottom',**kw)) 
                    for text,color in zip(list_of_strings,list_of_colors) ]
        xbox = HPacker(children=boxes,align="center",pad=0, sep=5)
        anchored_xbox = AnchoredOffsetbox(loc=3, child=xbox, pad=anchorpad,frameon=False,bbox_to_anchor=(0.2, -0.09),
                                          bbox_transform=ax.transAxes, borderpad=0.)
        ax.add_artist(anchored_xbox)

    # y-axis label
    if axis=='y' or axis=='both':
        boxes = [TextArea(text, textprops=dict(color=color, ha='left',va='bottom',rotation=90,**kw)) 
                     for text,color in zip(list_of_strings[::-1],list_of_colors) ]
        ybox = VPacker(children=boxes,align="center", pad=0, sep=5)
        anchored_ybox = AnchoredOffsetbox(loc=3, child=ybox, pad=anchorpad, frameon=False, bbox_to_anchor=(-0.10, 0.2), 
                                          bbox_transform=ax.transAxes, borderpad=0.)
        ax.add_artist(anchored_ybox)

def makeSummary(debug, sitename, titles, filename):

  years = debug[0][0][0]
  fig, axs = plt.subplots(len(debug),len(debug[0]), figsize=(14, 5*len(debug)), squeeze=False)
  for j, row in enumerate(axs):
    
    for i, ax in enumerate(row):
      site = sitename[j]
      ax.axhline(y=1.0, color='r', linewidth=0.3)

      ax.plot(years, np.array(debug[0][i][1])[:,2]/np.array(debug[0][i][1])[:,1], 
            color='b', linestyle='-', label='Storage')
      ax.plot(years, np.array(debug[0][i][1])[:,3]/np.array(debug[0][i][1])[:,1],
            color='b', linestyle='--', label='OOW Storage')
      computing_ratio = np.array(debug[0][i][1])[:,6]/np.array(debug[0][i][1])[:,5]
      ax.plot(years, computing_ratio, 
            color='orange', linestyle='-', label='Computing')
      ax.plot(years, np.array(debug[0][i][1])[:,7]/np.array(debug[0][i][1])[:,5],
            color='orange', linestyle='--', label='OOW Computing')
      sb = np.array(debug[0][i][1])[:,0]
      omsb = np.array(debug[0][i][1])[:,4]
      s = np.divide(sb, (sb+omsb), out=np.ones_like(sb), where=omsb!=0)
      ax.plot(years, s,
            color='green', linestyle='-', label='Storage fraction')
      ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
      ax.set_ylim([0,1.1*max(computing_ratio)])
      ax.set_title('Scenario {}'.format(i+1))
      multicolor_ylabel(ax,['Storage / Target', ' and ', 'Computing / Target'],['blue', 'black', 'orange'], 'y')
      legend_elements = [
        Line2D([0], [0], color='black', linestyle='-', label='Total'),
        Line2D([0], [0], color='black', linestyle='--', label='Out-of-warranty'),
        Line2D([0], [0], color='green', linestyle='-', label='Storage budget fraction')
      ]
      ax.legend(handles=legend_elements,
                loc='upper left', bbox_to_anchor=(0.5, 0.6),
                frameon=False)
      ax.text(0.5, 0.6, '{}, {}'.format(site, titles[i]),
              transform=ax.transAxes,
              horizontalalignment='left', verticalalignment='bottom')
  fig.savefig('drive/MyDrive/Facilities/Planning/summary/{}'.format(filename))
