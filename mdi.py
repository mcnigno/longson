from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, PatternFill, Color, Side, Alignment, NamedStyle


def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None, first_cell=None):
    """
    Apply styles to a range of cells as if they were a single cell.

    :param ws:  Excel worksheet instance
    :param range: An excel range to style (e.g. A1:F20)
    :param border: An openpyxl Border
    :param fill: An openpyxl PatternFill or GradientFill
    :param font: An openpyxl Font object
    """

    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)

    first_cell = first_cell
    if alignment:
        
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment

    rows = ws[cell_range]
    if font:
        first_cell.font = font

    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill



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
            'category_code': row[4].value,
            'title': row[3].value,
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
    cat_index[str(row[0].value)] = obj
    '''
    if row[2].value in cat_index:
        cat_index[row[0].value].append(obj)
    else:
        cat_index[str(row[0].value)] = [obj]
    '''

print('index len')
print('doc index len', len(doc_index))
print('cat index len', len(cat_index))
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
pdb_hash_cat = {}
only_pdb_doc = False
date_style = NamedStyle(name='datetime', number_format='DD/MM/YYYY')

for row in wps.iter_rows(min_row=2):
    
    if row[2].value[5:8] in pdb_hash_cat:
        pdb_hash_cat[row[2].value[5:8]] += 1 
    
    else:
        pdb_hash_cat[row[2].value[5:8]] = 1

    if row[2].value in pdb_hash:
        pdb_hash[row[2].value].append({
            'discipline': row[15].value,
            'document_no': row[2].value,
            'document_name': row[3].value,
            'revision': str(row[4].value),
            'transmittal': row[6].value,
            'trans_date': row[7].value,
            'require_action': row[5].value,
            'revised_plan': '',
            'issue_plan': '',
            'return_date': row[9].value,
            'owner_cmmt': row[8].value[:2]
        })
    else:
        pdb_hash[
            str(row[2].value)] = [{
                'discipline': row[15].value,
                'document_no': row[2].value,
                'document_name': row[3].value,
                'revision': str(row[4].value),
                'transmittal': row[6].value,
                'trans_date': row[7].value,
                'require_action': row[5].value,
                'revised_plan': '',
                'issue_plan': '',
                'return_date': row[9].value,
                'owner_cmmt': row[8].value[:2]
            }]
#print(pdb_hash)    
print('pdb_hash len', len(pdb_hash))
print('pdb_hash_cat len', len(pdb_hash_cat))

# Open the MDI Template
MDI_template = open('xls/template/MDR_Template.xlsx', mode='rb')
wmb = load_workbook(MDI_template, guess_types=True, data_only=True)
ws = wmb.active


start_row = 11
end_row = 19


fake_label = ['Purpose**','Rev.','Issue Plan','Revised Plan','Issue Actual', 'Transmittal no.', 'Return Date', 'Owner Cmmt*']


for cat, value in cat_index.items():
    
    cat_code = 'Category Code: ' + cat
    item = ''
    org = value['org']
    document_no = value['title']
    document_name = str(cat)
    classification = 'Class '+ str(value['class'])

    index = cat_index
    if only_pdb_doc:
        index = pdb_hash_cat

    if cat in index:
        _ = ws.cell(row=start_row, column=1, value=cat_code)
        ws.merge_cells(start_row=start_row, end_row=start_row, start_column=1, end_column=5)
        _.font = Font(b=True)
        _.fill = PatternFill("solid", fgColor="DDDDDD")
        _class = ws.cell(row=start_row, column=6)

        thin = Side(border_style="thin", color="000000")
        double = Side(border_style="double", color="ff0000")
        single = Side(border_style="medium", color="ff0000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="DDDDDD")
        font = Font(b=True, color="000000")
        al = Alignment(horizontal="left", vertical="center")
        
        cat_ranges = 'A'+ str(start_row)+':E'+str(start_row)
        #org_range = 'B'+ str(start_row+1)+':B'+str(start_row+8)
        style_range(ws, cat_ranges, border=border, fill=fill, font=font, alignment=al,first_cell=ws.cell(start_row,1))
        _class.fill = fill
        _class.border = border
        #org_range = 'B'+ str(start_row+1)+':B'+str(start_row+8)
        #style_range(ws, org_range, border=border, fill=fill, font=font, alignment=al,first_cell=ws.cell(start_row+1,2))
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        fill = PatternFill("solid", fgColor="DDDDDD")
        al = Alignment(horizontal="left", vertical="center")
        v_al = Alignment(horizontal="center", vertical="center")
        
        # Style The Category Header
        x = 7
        for n in range(13):
            n = ws.cell(row=start_row, column=x)
            n.border = border
            n.fill = fill
            x += 1

        
        # Set Catecory Code
        doc_counter = 0
        for doc, revisions in pdb_hash.items():
            if doc[5:8] == cat:
                doc_counter += 1
                
                
                
                # Set Document Info
                
                ws.cell(row=start_row+1, column=2, value=org)
                ws.cell(row=start_row+1, column=3, value=document_no)
                ws.cell(row=start_row+1, column=4, value=document_name)
                ws.cell(row=start_row+1, column=6, value=classification)
                
                #ws.merge_cells(start_column=2, end_column=2, start_row=start_row+1, end_row=start_row+8)
                #ws.merge_cells(start_column=3, end_column=3, start_row=start_row+1, end_row=start_row+8)
                #ws.merge_cells(start_column=4, end_column=4, start_row=start_row+1, end_row=start_row+8)

                
                # Set Data Label
                
                
                
                tmp_row = start_row + 1

                for label in fake_label:    
                    _ = ws.cell(row=tmp_row, column=7, value=label)
                    dotted = Side(border_style="dotted", color="000000")
                    _.border = Border(bottom=dotted)
                    if label == "Purpose**":
                        _.border = Border(top=thin, bottom=dotted)
                    elif label == 'Owner Cmmt*':
                        _.border = Border(bottom=thin)

                    
                    tmp_row += 1
                
                
                    
                tmp_col = 8
                tmp_row = start_row + 1
                for revision in revisions:
                    ws.cell(row=start_row+1, column=3, value=revision['document_no'])
                    ws.cell(row=start_row+1, column=4, value=revision['document_name'])
                    
                    purpose = ws.cell(row=tmp_row, column=tmp_col, value=revision['require_action'])
                    rev = ws.cell(row=tmp_row+1, column=tmp_col, value=revision['revision'])

                    issue_plan = ws.cell(row=tmp_row+2, column=tmp_col, value=revision['issue_plan'])
                    revised_plan = ws.cell(row=tmp_row+3, column=tmp_col, value=revision['revised_plan'])
                    issue_actual = ws.cell(row=tmp_row+4, column=tmp_col, value=revision['trans_date'])

                    trans = ws.cell(row=tmp_row+5, column=tmp_col, value=revision['transmittal'])

                    return_date = ws.cell(row=tmp_row+6, column=tmp_col, value=revision['return_date'])
                    owner_cmmt = ws.cell(row=tmp_row+7, column=tmp_col, value=revision['owner_cmmt'])


                    tmp_col += 1

                    dotted = Side(border_style="dotted", color="000000")
                    border = Border(bottom=dotted, right=thin)
                    
                    al = Alignment(horizontal='center')
                    purpose.border = Border(bottom=dotted,right=thin,top=thin)
                    purpose.alignment = al
                    rev.border = border
                    rev.alignment = al
                    issue_plan.border = border
                    issue_plan.alignment = al
                    issue_plan.number_format = 'DD/MM/YYYY'
                    revised_plan.border = border
                    revised_plan.alignment = al
                    revised_plan.number_format = 'DD/MM/YYYY'
                    issue_actual.border = border
                    issue_actual.alignment = al
                    issue_actual.number_format = 'DD/MM/YYYY'
                    trans.border = border
                    trans.alignment = al
                    return_date.border = border
                    return_date.alignment = al
                    return_date.number_format = 'DD/MM/YYYY'
                    owner_cmmt.border = Border(bottom=thin,right=thin)
                    owner_cmmt.alignment = al
                    
                thin = Side(border_style="thin", color="000000")
                border = Border(top=thin, left=thin, right=thin, bottom=thin)
                fill = PatternFill("solid", fgColor="DDDDDD")
                al = Alignment(horizontal="left", vertical="center")
                v_al = Alignment(horizontal="center", vertical="center")
                
                item_range = 'A'+ str(tmp_row)+':A'+str(tmp_row+7)
                org_range = 'B'+ str(tmp_row)+':B'+str(tmp_row+7)
                doc_no_range = 'C'+ str(tmp_row)+':C'+str(tmp_row+7)
                doc_name_range = 'D'+ str(tmp_row)+':E'+str(tmp_row+7)
                doc_class_range = 'F'+ str(tmp_row)+':F'+str(tmp_row+7)
                
                '''
                style_range(ws, item_range, border=border, alignment=al,first_cell=ws.cell(start_row+1,1))
                style_range(ws, org_range, border=border, alignment=v_al,first_cell=ws.cell(start_row+1,2))
                style_range(ws, doc_no_range, border=border, alignment=v_al, first_cell=ws.cell(start_row+1,3))
                style_range(ws, doc_name_range, border=border, alignment=al, first_cell=ws.cell(start_row+1,4))
                style_range(ws, doc_class_range, border=border, alignment=v_al,first_cell=ws.cell(start_row+1,6))
                '''

                
                
                start_row += 8
        
        start_row += 1



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


