"""
@author: graygoolsby
08/07/2018
Create lite Manufacturing Resource Planning system that outputs production runs by cross-referencing
inventory levels, sales orders, and WCL output
"""
import numpy as np
import pandas as pd

def mrp(month, day, year):
    filename = 'Paint_Production'+'_'+month+'.'+day+'.'+year
    stock = readFile('Paint_August_MTS_MTO')
    batch = readFile('Paint_August_Batch_SIzes')
    inventory = readFile('Paint_August_Inventory')
    reorder = readFile('Paint_August_Reorder_Qty')
    so = readFile('Paint_8.7.2018_Sales_Orders')
    
    production = checkStockItems(inventory, reorder, stock, batch)
    
    MTO = checkSalesOrders(inventory, batch, so)
    
    for run in MTO:
        production.append(run)
    
    production = pd.DataFrame(production)
    production.columns = ('Item', 'Batch Size', 'Status')
    
    writeFile(filename, production) 

# takes inventory, finds MTS items, compares inventory to reorder qty, and outputs production runs for necessary items
def checkStockItems(inventory, reorder, stock, batch):
    print('creating production runs for MTS items...')
    print("\n")
    
    """ write guts to make comparisions"""
    
    productionRuns = [['', 'MTS', '']]
    item1 = ['Item 1', '1800', 'Stock']
    item2 = ['Item 2','400','Stock-1']
    
    productionRuns.append(item1)
    productionRuns.append(item2)
    
    return productionRuns

# creates production runs for MTO items based on outstanding sales orders
def checkSalesOrders(inventory, batch, so):
    print ('creating production runs for MTO items...')
    print ("\n")
    
    """ write guts to make comparisons"""
    
    productionRuns = [['', 'MTO', '']]
    item3 = ['Item 3', '100', '8/15/2018' ]
    item4 = ['Item 4', '400', '8/20/2018']
    
    productionRuns.append(item3)
    productionRuns.append(item4)
    
    return productionRuns

# reads and returns data in file as data frame
# file names: 'Paint August MTS MTO', 'Paint August Batch Sizes', 'Paint August Inventory', 'Paint August Reorder Qty'
def readFile(filename):
    filename += '.xlsx'
    file = pd.ExcelFile(filename)
    sheets = file.sheet_names
    data = file.parse(sheets[0])
    print(filename +' read')
    print("\n")
    return data

# writes dataframe into excel file
def writeFile(filename, data):
    filename += '.xlsx'
    writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
    data.to_excel(writer, 'Sheet 1')
    writer.save()
    print(filename + 'created')
    print("\n")