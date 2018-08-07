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
#    so = readFile('Sales Orders')
    
    production = checkStockItems(inventory, reorder, stock, batch)
    
    MTO = (checkSalesOrders(inventory, batch))
    
    for run in MTO:
        production.append(run)
    
    production = pd.DataFrame(production)
    production.columns = ('Item', 'Batch Size', 'Status')
    
    
    
    
    
#    stockItem = stock['Item']
#    stockStatus = stock['Status']
#    batchItem = batch['Item']
#    batchSize = batch['Size']
#    inventoryItem = inventory['Item']
#    inventoryAmt = inventory['Qty']
#    reorderItem = reorder['Item']
#    reorderQty = reorder['Qty']
#    soItem = so['Item']
#    soQty = so['Qty']
#    print(stock)  headers : Item, Status
#    print("\n")
#    print(batch)  headers: Clean SKU, Batch Size
#    print("\n")    
#    print(inventory) headers: Item ID, Item, Gallons On Hand - Month
#    print("\n")
#    print(reorder)    
#    print("\n")
    
    
    
    writeFile(filename, production)

# takes inventory, finds MTS items, compares inventory to reorder qty, and outputs production runs for necessary items
def checkStockItems(inventory, reorder, stock, batch):
    print('creating production runs for MTS items...')
    print("\n")
    
    production = [['', 'MTS', '']]
    item1 = ['Item 1', '1800', 'Stock']
    item2 = ['Item 2','400','Stock-1']
    
    production.append(item1)
    production.append(item2)
    
    return production

# creates production runs for MTO items based on outstanding sales orders
def checkSalesOrders(inventory, batch):
    print ('creating production runs for MTO items...')
    print ("\n")
    
    production = [['', 'MTO', '']]
    item3 = ['Item 3', '100', '8/15/2018' ]
    item4 = ['Item 4', '400', '8/20/2018']
    
    production.append(item3)
    production.append(item4)
    
    return production

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