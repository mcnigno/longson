from openpyxl import Workbook, load_workbook

# Opent the MDI List
MDI_list = open('xls/lists/MDI_list.xlsx', mode='rb')
wlb = load_workbook(MDI_list, guess_types=True, data_only=True)
wls = wlb.active

category_list = set()
document_list = set() 

doc_index = {}
cat_index = {}

for row in wls.iter_rows(min_row=2):
    #print(row[0].value)
    #category.add(wls.cell(row, 1).value)

    category_list.add(row[0].value)
    document_list.add(row[2].value)
    
    obj = {
            'title': row[3].value,
            'category_code': row[4].value,
            'org': row[5].value,
            'cat_class': row[6].value,
            'class': row[7].value
            }


    # DOC INDEX
    if row[2].value in doc_index:
        doc_index[row[2].value].append(obj)
    else:
        doc_index[
            str(row[2].value)] = [obj]

    # CAT INDEX
    if row[2].value in cat_index:
        cat_index[row[0].value].append(obj)
    else:
        cat_index[
            str(row[0].value)] = [obj]

print('doc_index')
print(doc_index)
'''
category.add({
    'category_code': row[0].value,
    'document_id' : row[2].value,
    'org' : row[5].value,
    'category_class' : row[6].value
})
'''

print('there are', len(category_list), 'category and', len(document_list),'documents')

# Open the PDB List
PDB_list = open('xls/lists/PDB_list.xlsx', mode='rb')
wpb = load_workbook(PDB_list, guess_types=True, data_only=True)
wps = wpb.active

'''
for row in wls.iter_rows(min_row=2):
    doc = row[2].value
'''

'''
pdb_hash = []

for row in wps.iter_rows(min_row=2):
    pdb_hash.append({
        str(row[2].value) : {
            'revision': row[4].value,
            'transmittal': row[6].value,
            'tans. date': row[7].value,
            'require_action': row[12].value
        }
    })
print('pdb_hash len', len(pdb_hash))
'''

pdb_hash = {}

for row in wps.iter_rows(min_row=2):
    if row[2].value in pdb_hash:
        pdb_hash[row[2].value].append({
            'revision': row[4].value,
            'transmittal': row[6].value,
            'tans. date': row[7].value,
            'require_action': row[12].value
        })
    else:
        pdb_hash[
            str(row[2].value)] = [{
                'revision': row[4].value,
                'transmittal': row[6].value,
                'tans. date': row[7].value,
                'require_action': row[12].value
            }]
print(pdb_hash)    
print('pdb_hash len', len(pdb_hash))




'''
for cat in category_list:
    
    for doc in document_list:
        if doc in pdb_hash:
            print('Document', doc)
            print(pdb_hash[doc])

        for doc_id in pdb_hash:
            
            if doc in doc_id:
                print('Category:', cat)
                print('Document', doc)
                print('Revisions:')
                print('Revision', doc_id)
'''


 

# Open the MDI Template
MDI_template = open('xls/template/MDR_Template.xlsx', mode='rb')
wmb = load_workbook(MDI_template, guess_types=True, data_only=True)
ws = wmb.active


start_row = 11
end_row = 19

fake_category = ['G01','G02', 'G03', 'G04']
fake_label = ['Purpose**','Rev.','Issue Plan','Revised Plan','Issue Actual', 'Transmittal no.', 'Return Date', 'Owner Cmmt*']
fake_revisions = [['A1','B1','C1','D1','E1','F1','G1','H1'],['A2','B2','C2','D2','E2','F2','G2','H2']]


for cat in cat_index:
    for doc in pdb_hash:
        #print('CAT:', cat,'DOC:', doc[5:8])
        if doc[5:8] == cat:
            print('Documents for this CAT:', cat)
            print(doc)



for cat, value in cat_index.items():

    print(cat, value)
    # Set Catecory Code
    
    cat_code = 'Category Code: ' + cat
    item = ''
    org = value['org']
    document_no = value['title']
    document_name = str(cat)
    classification = value['class']

    ws.cell(row=start_row, column=1, value=cat_code)
    ws.merge_cells(start_row=start_row, end_row=start_row, start_column=1, end_column=5)
    
    # Set Document Info
    
    ws.cell(row=start_row+1, column=2, value=org)
    ws.cell(row=start_row+1, column=3, value=document_no)
    ws.cell(row=start_row+1, column=4, value=document_name)
    ws.cell(row=start_row+1, column=6, value=classification)
    
    # Set Data Label
    tmp_row = start_row + 1

    for label in fake_label:    
        ws.cell(row=tmp_row, column=7, value=label)
      
        tmp_row += 1
    
      
        
    tmp_col = 8
    for revision in fake_revisions:
        for data in revision:
            ws.cell(row=tmp_row, column=tmp_col, value=cat+data)
        tmp_col += 1
        
    start_row += 9








wmb.save('xls/MDI_TEST.xlsx')





'''
startCol = 7
startRow = 
endCol = 
endRow = 
sheetReceiving = 
copiedData = 


def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected
         

#Paste range
#Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1,1):
        countCol = 0
        for j in range(startCol,endCol+1,1):
            
            sheetReceiving.cell(row = i, column = j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1
def createData():
    print("Processing...")
    selectedRange = copyRange(1,2,4,14,sheet) #Change the 4 number values
    pastingRange = pasteRange(1,3,4,15,temp_sheet,selectedRange) #Change the 4 number values
    #You can save the template as another file to create a new file here too.s
    template.save("xls/MDI_TEST.xlsx")
    print("Range copied and pasted!")
'''
'''
for row in range(1,9):
    #doc_block = ws[start_row:end_row]

    ws.append(['x','y','z'])
    start_row +1

'''


