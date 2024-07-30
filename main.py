# import modules
import pandas as pd
import requests
pd.options.mode.chained_assignment = None  # default='warn'

 
#function to check if url is working
# pass the url
def url_ok(url):
     
    # exception block
    try:
       
        # pass the url into
        # request.hear
        response = requests.head(url)
         
        # check the status code
        if response.status_code == 200:
            return True
        else:
            return False
    except requests.ConnectionError as e:
        return e


#sets up the dataframe and other variables
df = pd.read_excel('./Excel-Sheets/ogSku.xlsx')
dictDataList = []
dictDataListLens = []
baseLink = 'https://www.beyondcell.com/0068c3WQxX_retail_images/OpticGard/Opticgard%20NO%20Logo/'

#gets the list of skus
skus = df['SKU'].tolist()
isLens = False

#iterates through the sku
for index, skuname in enumerate(skus):
    allSkuPic = []
    dictData = {}
    
    #if it is a lens remove everything until the _2
    if '_2' in skuname:
        parent = skuname[:skuname.index('_2')]
        isLens = True

    #removes the last two numbers if there isnt a DE
    elif 'DE' not in skuname:
        parent = skuname[:-2]

    #cuts off the SKU name right at the start of DE
    elif 'DE' in skuname:
        parent = skuname[:skuname.index('DE')]
    

    #makes the correct parent for the eps carry
    parent = parent.replace('OGLCHEPSCAR', 'OGLCHEPSCARRY').replace('OGLC', 'OG')
    skuname = skuname.replace('OGLCHEPSCAR', 'OGLCHEPSCARRY')

    #creates the link name
    linkName = baseLink + parent + '/' + skuname 

    #adds sku to dict
    dictData['SKU'] = skuname

    #checks to see if sku has x2 link and adds it in
    if isLens and url_ok(linkName + '.jpg'):
        dictData['X2 Links'] = linkName + '.jpg'

    elif url_ok(linkName + 'x2.jpg'):
        dictData['X2 Links'] = linkName + 'x2.jpg'

    elif url_ok(linkName + '.jpg'):
        dictData['X2 Links'] = linkName + '.jpg'
        
    else:
        dictData['X2 Links'] = 'Something went wrong ask Darien'
    

    #adds the all pic links so it is right after the x2
    dictData['All Pic Links Combined'] = ''

    #checks the x links to see if it works and if it does appends it to the dataframe
    if not isLens:
        for i in range(3, 13):
            urlJpg = linkName + 'x' + str(i) + '.jpg'
            urlPng = linkName + 'x' + str(i) + '.png'
            columnName = 'Picture Link X' + str(i)


            #checks to see if the jpg url works and changes it if it does
            if url_ok(urlJpg):
                dictData[columnName] = urlJpg
                allSkuPic.append(urlJpg)
            
            #same as jpg but png
            elif url_ok(urlPng):            
                dictData[columnName] = urlPng
                allSkuPic.append(urlPng)
            else:
                dictData[columnName] = ''
        
    #adds the list to the all pic links column
    dictData['All Pic Links Combined'] = ', '.join(allSkuPic)

    #appends dictData to correct list
    if isLens:
        dictDataListLens.append(dictData)
        isLens = False
    else:
        dictDataList.append(dictData)
    
    #prints the percent of amount done
    print(f'Percent Done: {(index + 1)/len(skus) * 100:.2f}%')

df = pd.DataFrame.from_dict(dictDataList)
dfLens = pd.DataFrame.from_dict(dictDataListLens)
    
writer = pd.ExcelWriter('ogPictureLink_output.xlsx', engine = 'openpyxl')
df.to_excel(writer, sheet_name = 'Optic Cover')
dfLens.to_excel(writer, sheet_name = 'Lens')
writer.close()