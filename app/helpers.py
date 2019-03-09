from .models import Janus, Doc_list, Pdb, Category
from app import db
from datetime import datetime
import openpyxl
from flask import flash

janus_not_in_document_list = []
pdb_not_in_document_list = []

def janus_upload():
    session = db.session
    janus_file = open('./liste/janus.txt',mode='rb')
    
    row_number = 0

    doc_list = [x.client_reference for x in session.query(Doc_list).all()]
    
    for line in janus_file:
        
        #Skip in Headers at line 1
        row_number += 1
        
        
        if row_number == 1:
            continue
        
        # Split Row by |
        row = str(line).split('|')
        
        # Create Janus row object
        
        janus = Janus(
    
            linenumber = row[0],
            cat = row[1],
            doc_reference = row[2],
            weight = row[9],
            title = row[11],
            milestone_chain = row[12],
            dbs = row[13],
            wbs = row[14],
            fcr_one = row[15],
            fcr_two = row[16],
            fcr_three = row[17],
            mscode = row[19],
            cumulative = row[20],
            obs = row[23],
            initial_plan_date = date_parse(row[24]),
            revised_plan_date = date_parse(row[25]),
            forecast_date = date_parse(row[26]),
            actual_date = date_parse(row[27])
        )
    
        janus.created_by_fk = '1'
        janus.client_reference_id = row[8]

        # Check if the janus doc is or not in document list
        
        if row[8] in doc_list:
            session.add(janus)
        else:
            janus.client_reference_id = 'wrong_client_ref'
            session.add(janus)
            janus_not_in_document_list.append(janus)
    
    session.commit()
    
    print('Janus document not in document list')
    print(janus_not_in_document_list)
    return 'Janus updated!'

def date_parse(date):
    
    if isinstance(date, datetime): return date
    if date == '' or date is None: return None

    return datetime.strptime(date,'%d-%b-%y')
    
def document_list_upload():
    session = db.session
    doclist = openpyxl.load_workbook('./liste/doclist.xlsx')
    doclist_ws = doclist.active
    count_id = 0
    wrong_ref =  session.query(Doc_list).filter(Doc_list.client_reference == 'wrong_client_ref').first()
    empty_client_references =  session.query(Doc_list).filter(Doc_list.client_reference == 'EmptyClientReferences').first()            
    for row in doclist_ws.iter_rows(min_row=2):
        
        doclist_row = Doc_list(
            cat = row[0].value,
            title = row[3].value,
            org = row[5].value,
            cat_class = row[6].value,
            class_two = row[7].value,
            weight = row[8].value,
            doc_reference = row[12].value
        )
        doclist_row.created_by_fk = '1'
        if row[2].value is None:
            count_id += 1
            doclist_row.client_reference = 'EmptyClientReferences' + str(count_id)
        else:
            doclist_row.client_reference = row[2].value
        session.add(doclist_row)
        # Check if Wrong References Doc for Janus exist, otherwise add it

        
        if wrong_ref is None:
            wrong_ref = Doc_list()
            wrong_ref.created_by_fk = '1'
            wrong_ref.client_reference = 'wrong_client_ref'
            session.add(wrong_ref)
    session.commit()
    return 'Document List updated!'
    
def pdb_list_upload():
    session = db.session
    pdblist = openpyxl.load_workbook('./liste/pdb.xlsx')
    pdblist_ws = pdblist.active
 
    doc_list = [x.client_reference for x in session.query(Doc_list).all()]
    
                
    for row in pdblist_ws.iter_rows(min_row=2):
        pdb = Pdb(
            doc_reference = row[0].value,
            title = row[1].value,
            revision_number = row[2].value,
            revision_date = date_parse(row[3].value),
            document_revision_object = row[4].value,
            
            discipline = row[6].value,
            transmittal_date = date_parse(row[7].value),
            transmittal_reference = row[8].value,
            specific_transmittal_number = row[9].value,
            required_action = row[10].value,
            response_due_date = date_parse(row[11].value),
            actual_response_date = date_parse(row[12].value),
            document_status = row[13].value,
            client_transmittal_ref_number = row[14].value,
            remarks = row[15].value,
        )
        pdb.created_by_fk = '1'
        pdb.client_reference_id = row[5].value
        

        # Check if the pdb doc is or not in document list
        '''
        if session.query(Doc_list).filter(Doc_list.client_reference == row[5].value).first():
            session.add(pdb)
        '''
        if row[5].value in doc_list:
            session.add(pdb)
        else:
            pdb.client_reference_id = 'wrong_client_ref'
            session.add(pdb)
            pdb_not_in_document_list.append(pdb)

    session.commit()
    return 'PDB updated!'
    #print('Pdb document not in document list')
    #print(pdb_not_in_document_list)

# Open the Category Code List
def category_upload():
    CAT_List = open('./liste/Category_List.xlsx', mode='rb')
    wcb = openpyxl.load_workbook(CAT_List, data_only=True)

    sheet_list = wcb.sheetnames
    print('Sheets in Category Code List')

    cat_description = {}
    cat_duplicate =[]
    session = db.session
    for name in sheet_list[5:]:
        wcs = wcb[name]
        #print(name)
        for row in wcs.iter_rows(min_row=9):
            #print(row[0].value,row[1].value,row[2].value)
            
            if row[1].value is not None:
                cat = Category(
                    sheet_name = name,
                    code = row[0].value,
                    information = row[1].value,
                    description = row[2].value,

                )
                session.add(cat)
    try:
        session.commit()
    except:
        flash('Something Went Wrong: Check you Category file! Row starts at line 9.')

    return 'category done!'
    '''
                print(name ,row[0].value, row[1].value)
                if row[0].value not in cat_description:
                    cat_description[row[0].value] = row[1].value
                else: 
                    print(row[0].value, 'is a duplicate??')
                    obj = {
                        'SheetName': name,
                        'Code': row[0].value,
                        'Information': row[1].value,
                        'Description': row[2].value,
                        'Duplicate': cat_description[row[0].value]
                    }
                    cat_duplicate.append(obj)
        print('')        
        print('the cat count is:', cat_count)
        print('the len of category codes is:', len(cat_description))
        print('') 
        print('List of duplicate codes in category:', cat_duplicate)
        print('') 
    '''

def update_all():
    session = db.session
    
    deleted_pdb = session.query(Pdb).delete()
    
    deleted_janus = session.query(Janus).delete()
    deleted_cat = session.query(Category).delete()
    deleted_doc = session.query(Doc_list).delete()
    
    session.commit()

    doc = document_list_upload()
    print('Deleted',deleted_pdb, 'rows from pdb, row now in pdb:', len(session.query(Pdb).all()))
    pdb = pdb_list_upload()
    print('Uploaded pdb, number of row', len(session.query(Pdb).all()))

    janus = janus_upload()
    cat = category_upload()
    

    
    return doc, pdb, janus, cat

def check_pdb_not_in_janus():
    print('Check PDB Vs Janus')
    session = db.session
    #document_list = session.query(Doc_list).all()
    janus_list = session.query(Janus).all()
    pdb_list = session.query(Pdb).all()
    janus_ref = [x.client_reference for x in janus_list]

    for doc in pdb_list:
        #print(doc.client_reference)
        if doc.client_reference not in janus_ref:
            print('pdb and janus match', doc.client_reference, doc.revision_number, doc.client_reference)
    
    return 