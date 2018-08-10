"""
@author: graygoolsby, ICP Summer Operation Intern
August 2018
Create lite Manufacturing Resource Planning system that outputs production runs by cross-referencing
inventory levels, sales orders, MTS/MTO designation, restock quantities, and WCL output
"""
import pandas as pd

# master function that runs full MRP process. CALL THIS FUNCTION
def mrp(prodLine, month, day, year):
    # error proof input
    prodLine = str(prodLine)
    month = str(month)
    day = str(day)
    year =str(year)
    # create filename to write output to
    filename = prodLine+'_Production'+'_'+month+'.'+day+'.'+year+'.xlsx'

    # read files for MRP
    stock = readFile(prodLine+'_'+month+'_MTS_MTO_Tolled').set_index('Item').T.to_dict('list')
    batch = readFile(prodLine+'_'+month+'_Batch_SIzes').set_index('Item').T.to_dict('list')
    inventory = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Inventory').set_index('Item').T.to_dict('list')
    reorder = readFile(prodLine+'_'+month+'_Reorder_Qty').set_index('Item').T.to_dict('list')
    so = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Sales_Orders').values.tolist()
    schedule = readFile(prodLine+'_'+month+'.'+day+'.'+year+'_Scheduled_Production').values.tolist()

    # perform MRP and create production and purchase needs
    so = cleanSO(so)
    mts = checkStockItems(inventory, reorder, stock, batch)
    itemRuns = checkSalesOrders(inventory, batch, so, mts, stock)
    productionRuns = removeScheduledBatches(schedule, itemRuns)
    production = identifySplitFills(productionRuns)

    # turn output into format to write to excel
    production = pd.DataFrame(production)
    production.columns = ('Item', 'Batch Size (Gal)', 'Status', 'Qty', 'Split Fill')

    # write final output to excel
    writeFile(filename, production) 

# takes inventory, finds MTS items, compares inventory to reorder qty, and outputs production runs for necessary items
def checkStockItems(inventory, reorder, stock, batch):
    print('creating production runs for MTS items...')
    print("\n")
    # create header for output and objects to store data for production and data gaps
    productionRuns = [['MTS', '', '', 'Current Stock']]
    items = []
    issue = []

    # look at every item in the inventory 'file'
    for item in inventory:
        # object to store data of individual production runs
        run = []

        # prevents errors that arise from data gaps
        if(item in reorder and item in stock):

            # check inventory levels against reorder points
            if((float(reorder[item][0])*.75)>=float(inventory[item][0])):

                # only create production in this section for stocked items
                if(stock[item][0] == 'TTS'):
                    run = [item, '', 'Buy-1', float(inventory[item][0])]
                if(stock[item][0] == 'MTS'):
                    run = [item, '', 'Stock-1', float(inventory[item][0])]

                # filters out elements that represent items that don't need to be restocked
                if(len(run)>0):
                    items.append(run)
    
            if(float(reorder[item][0])>=float(inventory[item][0]) and float((reorder[item][0])*.75) < float(inventory[item][0])):

                if(stock[item][0] == 'TTS'):
                    run = [item, '', 'Buy', float(inventory[item][0])]
                if(stock[item][0] == 'MTS'):
                    run = [item, '', 'Stock', float(inventory[item][0])]

                if(len(run)>0):
                    items.append(run)
        # catch items that have data gaps
        else:
            if(inventory[item][0]>0):
                issue.append([item,inventory[item][0]])

    # add batch sizes to items that are made in-house
    for item in items:
        if(item[2] != 'Buy' and item[2] != 'Buy-1'):
            cleanSKU= item[0][:item[0].find('-')]
            item[1] = batch[cleanSKU][0]

    # add final production infomation to production list
    for item in items:
        productionRuns.append(item)

    # return items with data gaps to user
    print('check data on these items (in inventory but no reorder QTY or stock designation):')
    print(issue)
    print("\n")

    return productionRuns

# creates production runs for MTO items based on outstanding sales orders
def checkSalesOrders(inventory, batch, so, production, stock):
    print ('creating production runs for MTO items...')
    print ("\n")

    # create header for MTO section of output and create objects to store production and data gaps
    productionRuns = [['MTO', 'Batch Size', 'Order due', 'Needed for orders' ]]
    items = []
    issue = []

    # look at every item in Sales Order 'file'
    for item in so:
        # object to store production information
        run = []

        # filter out items that have been addressed in MTS section
        if(item[0] in stock and stock[item[0]][0] == 'MTS'):
            pass
        if(item[0] in stock and stock[item[0]][0] == 'TTS'):
            pass

        # find items do not have sufficient inventory to cover sales order
        if(item [0] in stock and item[0] in inventory and float(inventory[item[0]][0])<float(item[1])):

            # add tags for buying or making products and ensure no MTS or TTS items in MTO section
            if(stock[item[0]][0] == 'TTO'):
                run = [item[0], '', str(stock[item[0]])+' '+item[2], float(item[1])-float(inventory[item[0]][0])]
            if(stock[item[0]][0] == 'TTB'):
                run = [item[0], '', str(stock[item[0]])+' '+item[2], float(item[1])-float(inventory[item[0]][0])]
            if(stock[item[0]][0] == 'MTO'):
                run = [item[0], '', item[2], float(item[1])-float(inventory[item[0]][0])]

            # filters out objects that represent items that have sufficient inventory
            if(len(run)>0):
                items.append(run)

        # catch all items in sales orders that don't currently have inventory
        if(item[0] in stock and item[0] not in inventory):

            # add tags for buying or making products and ensure no MTS or TTS items in MTO section            
            if(stock[item[0]][0] == 'TTO' ):
                run = [item[0], '', 'Buy '+str(item[2]), float(item[1])]
            if(stock[item[0]][0] == 'TTB'):
                run = [item[0], '', 'Buy '+str(item[2]), float(item[1])]                
            if(stock[item[0]][0] == 'MTS'):
                run = [item[0], '', item[2], float(item[1])]

            # filters out unnecessary objects         POSSIBLY REDUNDENT
            if(len(run)>0):
                items.append(run)

        # catch all items that have data gaps
        else:
            issue.append(item[0])

    # return items with data gaps to user
    print('check data on these items (ordered but no stock designation):')
    print(issue)
    print("\n")

    # add batches to items that are made in house
    for item in items:
        if(not(item[2][:3] == 'Buy')):
            cleanSKU = item[0][:item[0].find('-')]
            if(cleanSKU in batch):
                item[1] = batch[cleanSKU][0]
            else:
                item[1] = 100

    # add production info to MTO section      
    for item in items:
        productionRuns.append(item)

    # add MTO section under MTS section    
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
    # counter to help tag MTO items for multiple batches
    batches = 0

    # look at all items that need to be made
    for item in production:

        # skip MTS and MTO lines (no need to tag these)
        if(item[0] == 'MTS' or item[0] == 'MTO'):
            item.append('')
            batches += 1

        # tag all other lines
        else:

            # if in MTO section
            if(batches > 1):

                # if multiple batches are needed to cover sales order qty
                if(item[3]/item[1] > 1):
                    item.append('MULTIPLE BATCHES')

            # translate item into key for batch sizes
            cleanSKU = item[0][:item[0].find('-')]
            # counter for split fills
            count = 0

            # look at every other item that needs to be produced
            for item2 in production:
                # translate other items into key for batch sizes
                cleanSKU2 = item2[0][:item2[0].find('-')]
                # if they are same item in different sizes
                if(cleanSKU==cleanSKU2):
                    count+=1

            # there is more that one item from same product and it's made in-house
            if(count>1 and item[2][:3] != 'Buy' ):

                # check if MULTIPLE BATCHES tag is there, and add 'Split Fill' tag
                if(len(item)<5):
                    item.append('Split Fill')

                else:
                    item[4] = item[4]+ ' - Split FIll'

    return production

# look at current batch schedule and remove all from production
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
