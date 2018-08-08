"""
@author: graygoolsby, ICP Summer Operation Intern
August 2018
Create lite Manufacturing Resource Planning system that outputs production runs by cross-referencing
inventory levels, sales orders, MTS/MTO designation, restock quantities, and WCL output
"""
import numpy as np
import pandas as pd

def mrp(month, day, year):
    filename = 'Paint_Production'+'_'+month+'.'+day+'.'+year+'.xlsx'
    stock = readFile('Paint_August_MTS_MTO').set_index('Item').T.to_dict('list')
    batch = readFile('Paint_August_Batch_SIzes').set_index('Item').T.to_dict('list')
    inventory = readFile('Paint_August_Inventory').set_index('Item').T.to_dict('list')
    reorder = readFile('Paint_August_Reorder_Qty').set_index('Item').T.to_dict('list')
    so = readFile('Paint_8.7.2018_Sales_Orders').values.tolist()
    
    so = cleanSO(so)
    
    production = checkStockItems(inventory, reorder, stock, batch)
    
    MTO = checkSalesOrders(inventory, batch, so)
    
    for run in MTO:
        production.append(run)
    
    production = identifySplitFills(production)
    
    production = pd.DataFrame(production)
    production.columns = ('Item', 'Batch Size', 'Status', 'Gallons', 'Split Fill')
    
    writeFile(filename, production) 

#    print(stock)  #headers : Item, Status
#    print("\n")
#    print(batch) # headers: Item, Size
#    print("\n")    
#    print(inventory) #headers: Item, Qty
#    print("\n")
#    print(reorder) #headers: Item, Qty
#    print("\n")
#    print(so) #headers: Item, Qty, Ship Date

# takes inventory, finds MTS items, compares inventory to reorder qty, and outputs production runs for necessary items
def checkStockItems(inventory, reorder, stock, batch):
    print('creating production runs for MTS items...')
    print("\n")
    
    """ add in column for qty of some sort """
    productionRuns = [['MTS', '', '', 'Current Stock']]
    items = []
    for item in inventory:
        run = []
        if(item in reorder and item in stock and stock[item][0] == 'MTS'):
            if(int(reorder[item][0])>=int(inventory[item][0])):
                run = [item, '', 'Stock-1', int(inventory[item][0])]
                items.append(run)
            elif((int(reorder[item][0])*1.1)>=int(inventory[item][0])):
                run = [item, '', 'Stock', int(inventory[item][0])]
                items.append(run)

    for item in items:
        cleanSKU= item[0][:item[0].find('-')]
        item[1] = batch[cleanSKU][0]

    for item in items:
        productionRuns.append(item)

    return productionRuns

# creates production runs for MTO items based on outstanding sales orders
def checkSalesOrders(inventory, batch, so):
    print ('creating production runs for MTO items...')
    print ("\n")
    
    """ write guts to make comparisons"""
    
    productionRuns = [['MTO', 'Batch Size', 'Status', 'Needed for orders' ]]
    item3 = ['Item 3', '8/15/2018' ,'100', 15] 
    item4 = ['Item 4', '8/20/2018', '400', 50]
    
    productionRuns.append(item3)
    productionRuns.append(item4)
    
    return productionRuns

def cleanSO(so):
    for item in so:
        item[2] = str(item[2])[:10]
    return so

def identifySplitFills(production):
    for item in production:
        if(item[0] == 'MTS' or item[0] == 'MTO'):
            pass
        else:
            cleanSKU = item[0][:item[0].find('-')]
            count = 0
            for item2 in production:
                cleanSKU2 = item2[0][:item2[0].find('-')]
                if(cleanSKU==cleanSKU2):
                    count+=1
            if(count>1):
                item.append('Split Fill')
    return production

# reads and returns data in file as data frame
# file names: 'Paint August MTS MTO', 'Paint August Batch Sizes', 'Paint August Inventory', 'Paint August Reorder Qty', 'Paint_DATE_Sale_Orders'
def readFile(filename):
    if(filename[filename.rfind('.'):] != 'xlsx'):
        filename += '.xlsx'
    file = pd.ExcelFile(filename)
    sheets = file.sheet_names
    data = file.parse(sheets[0])
    print(filename +' read')
    print("\n")
    return data

# writes dataframe into excel file
def writeFile(filename, data):
    writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
    data.to_excel(writer, 'Sheet 1')
    writer.save()
    print(filename + ' created')
