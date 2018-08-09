"""
@author: graygoolsby, ICP Summer Operation Intern
August 2018
Create lite Manufacturing Resource Planning system that outputs production runs by cross-referencing
inventory levels, sales orders, MTS/MTO designation, restock quantities, and WCL output
"""
import pandas as pd

# master function that runs full MRP process. CALL THIS FUNCTION
def mrp(prodLine, month, day, year):
    prodLine = str(prodLine)
    month = str(month)
    day = str(day)
    year =str(year)
    filename = prodLine+'_Production'+'_'+month+'.'+day+'.'+year+'.xlsx'
    
    stock = readFile(prodLine+'_'+month+'_MTS_MTO').set_index('Item').T.to_dict('list')
    batch = readFile(prodLine+'_'+month+'_Batch_SIzes').set_index('Item').T.to_dict('list')
    inventory = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Inventory').set_index('Item').T.to_dict('list')
    reorder = readFile(prodLine+'_'+month+'_Reorder_Qty').set_index('Item').T.to_dict('list')
    tolled = readFile(prodLine+'_'+month+'_Tolled_Items').set_index('Item').T.to_dict('list')
    so = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Sales_Orders').values.tolist()
    schedule = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Scheduled_Production').values.tolist()

    so = cleanSO(so)

    mts = checkStockItems(inventory, reorder, stock, batch, tolled)

    itemRuns = checkSalesOrders(inventory, batch, so, mts, tolled)

    productionRuns = removeScheduledBatches(schedule, itemRuns)

    production = identifySplitFills(productionRuns)

    production = pd.DataFrame(production)
    production.columns = ('Item', 'Batch Size', 'Status', 'Gallons', 'Split Fill')

    writeFile(filename, production) 

# takes inventory, finds MTS items, compares inventory to reorder qty, and outputs production runs for necessary items
def checkStockItems(inventory, reorder, stock, batch, tolled):
    print('creating production runs for MTS items...')
    print("\n")

    productionRuns = [['MTS', '', '', 'Current Stock']]
    items = []
    for item in inventory:
        run = []
        if(item in reorder and item in stock and stock[item][0] == 'MTS'):
            if((float(reorder[item][0])*.75)>=float(inventory[item][0])):
                if(item in tolled):
                    run = [item, '', 'Buy-1', float(inventory[item][0])]
                else:
                    run = [item, '', 'Stock-1', float(inventory[item][0])]
                items.append(run)
            elif(float(reorder[item][0])>=float(inventory[item][0])):
                if(item in tolled):
                    run = [item, '', 'Buy', float(inventory[item][0])]
                else:
                    run = [item, '', 'Stock', float(inventory[item][0])]
                items.append(run)

    for item in items:
        cleanSKU= item[0][:item[0].find('-')]
        item[1] = batch[cleanSKU][0]

    for item in items:
        productionRuns.append(item)

    return productionRuns

# creates production runs for MTO items based on outstanding sales orders
def checkSalesOrders(inventory, batch, so, production, tolled):
    print ('creating production runs for MTO items...')
    print ("\n")
    productionRuns = [['MTO', 'Batch Size', 'Order due', 'Needed for orders' ]]
    items = []
    for item in so:
        run = []
        if(item[0] in inventory and float(inventory[item[0]][0])<float(item[1])):
            if(item[0] in tolled):
                run = [item[0], '', tolled[item][0]+' '+item[2], float(item[1])-float(inventory[item[0]][0])]
            else:
                run = [item[0], '', item[2], float(item[1])-float(inventory[item[0]][0])]
            items.append(run)
        elif(item[0] not in inventory):
            if(item[0] in tolled):
                run = [item[0], '', tolled[item][0]+' '+item[2], float[item[1]]]
            else:
                run = [item[0], '', item[2], float(item[1])]
            items.append(run)
        else:
            pass
        
    for item in items:
        cleanSKU = item[0][:item[0].find('-')]
        if(cleanSKU in batch):
            item[1] = batch[cleanSKU][0]
        else:
            item[1] = 100
            
    for item in items:
        productionRuns.append(item)
        
    for run in productionRuns:
        production.append(run)

    return production

# cleans time from due dates on sales orders
def cleanSO(so):
    for item in so:
        item[2] = str(item[2])[:10]
    return so

# identifies and tags all products that are same item, just different size
def identifySplitFills(production):
    batches = 0
    for item in production:
        if(item[0] == 'MTS' or item[0] == 'MTO'):
            batches += 1
        else:
            if(batches > 1):
                if(item[3]/item[1] > 1):
                    item.append('MULTIPLE BATCHES')
            cleanSKU = item[0][:item[0].find('-')]
            count = 0
            for item2 in production:
                cleanSKU2 = item2[0][:item2[0].find('-')]
                if(cleanSKU==cleanSKU2):
                    count+=1
            if(count>1):
                if(len(item)<5):
                    item.append('Split Fill')
                else:
                    item[4] = item[4]+ ' - Split FIll'
    return production

def removeScheduledBatches(schedule, production):
    for run in production:
        sku = run[0]
        for batch in schedule:
            if(sku == batch[0]):
                production.remove(run)

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
