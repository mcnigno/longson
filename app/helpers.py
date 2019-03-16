from .models import Janus, Doc_list, Pdb, Category, SourceFiles, Sourcetype
from app import db
from datetime import datetime
import openpyxl
from flask import flash
from flask_appbuilder.models.sqla.interface import SQLAInterface
from config import UPLOAD_FOLDER

janus_not_in_document_list = []
pdb_not_in_document_list = []

def janus_upload(source):
    session = db.session
    janus_file = open(UPLOAD_FOLDER + source)
    row_number = 0
    count_janus = 0
    doc_list = [x.client_reference for x in session.query(Doc_list).all()]
    try:
        for line in janus_file:
            
            #Skip in Headers at line 1
            row_number += 1
            
            
            if row_number == 1:
                continue
            
            # Split Row by |
            row = str(line).split('|')
            
            # Create Janus row object
            #print('line_number:', row[0])
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
                count_janus += 1
            else:
                janus.client_reference_id = 'wrong_client_ref'
                session.add(janus)
                #janus_not_in_document_list.append(janus)

        session.commit()
        return str(count_janus) + ' Janus Updated'
    except:
        return 'Janus FAIL: check your source file'
    

def date_parse(date):
    
    if isinstance(date, datetime): return date
    if date == '' or date is None: return None

    return datetime.strptime(date,'%d-%b-%y')
    
def document_list_upload(source):
    print('document_list_upload***')
    session = db.session

    doclist = openpyxl.load_workbook(UPLOAD_FOLDER + source)
    doclist_ws = doclist.active
    count_id = 0
    count_doc = 0
    wrong_ref = session.query(Doc_list).filter(Doc_list.client_reference == 'wrong_client_ref').first()
    #empty_client_references = session.query(Doc_list).filter(Doc_list.client_reference == 'EmptyClientReferences').first()            
    try:    
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
            count_doc += 1

            
            # Check if Wrong References Doc for Janus exist, otherwise add it
            if wrong_ref is None:
                wrong_ref = Doc_list()
                wrong_ref.created_by_fk = '1'
                wrong_ref.client_reference = 'wrong_client_ref'
                session.add(wrong_ref)
        
        session.commit()
        
        return str(count_doc) + ' Document List updated!'
    except:
        return 'Document List FAIL: check your source file.'


def pdb_list_upload(source):
    session = db.session
    pdblist = openpyxl.load_workbook(UPLOAD_FOLDER + source)
    pdblist_ws = pdblist.active
 
    doc_list = [x.client_reference for x in session.query(Doc_list).all()]
    count_pdb = 0
    try:           
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
            
            if row[5].value in doc_list:
                session.add(pdb)
                count_pdb += 1
            else:
                pdb.ex_client_reference = pdb.client_reference_id
                pdb.client_reference_id = 'wrong_client_ref'
                session.add(pdb)
                pdb_not_in_document_list.append(pdb)

        session.commit()
        return str(count_pdb) + ' PDB updated!'
    except:
        return 'PDB FAIL: check your source file.'  

# Open the Category Code List
def category_upload(source):
    wcb = openpyxl.load_workbook(UPLOAD_FOLDER + source)
    sheet_list = wcb.sheetnames
    
    session = db.session
    count_cat = 0
    try:
        for name in sheet_list[5:]:
            wcs = wcb[name]
            
            for row in wcs.iter_rows(min_row=9):
                            
                if row[1].value is not None:
                    cat = Category(
                        sheet_name = name,
                        code = row[0].value,
                        information = row[1].value,
                        description = row[2].value,

                    )
                    session.add(cat)
                    count_cat += 1
        
        session.commit()
        return str(count_cat) + ' Categories updated!'
    except:
        return 'Categories FAIL: check your source file.'
    


def update_all():
    session = db.session
    
    deleted_pdb = session.query(Pdb).delete()
    deleted_janus = session.query(Janus).delete()
    deleted_cat = session.query(Category).delete()
    deleted_doc = session.query(Doc_list).delete()
    
    session.commit()

    sf = SourceFiles
    files = dict([(str(x.source_type),x.file_source) for x in session.query(sf).all()])

    doc = document_list_upload(files['Document List'])
    pdb = pdb_list_upload(files['PDB'])
    janus = janus_upload(files['Janus'])
    cat = category_upload(files['Categories'])
    

    
    return doc, pdb, janus, cat, deleted_doc, deleted_pdb, deleted_janus, deleted_cat

def check_pdb_not_in_janus():
    print('Check PDB Vs Janus')
    session = db.session
    #document_list = session.query(Doc_list).all()
    janus_list = session.query(Janus).all()
    pdb_list = session.query(Pdb).all()
    janus_ref = [x.doc_reference for x in janus_list]
    pdb_not_in_janus = []
    for doc in pdb_list:
        #print(doc.client_reference)
        if doc.doc_reference not in janus_ref:
            print('pdb not in janus', doc.client_reference, doc.revision_number, doc.doc_reference)
            pdb_not_in_janus.append(doc)
        return pdb_not_in_janus

def init_file_type():
    session = db.session
    exist_list = [x.source_type for x in db.session.query(Sourcetype).all()]
    type_list = ['Document List','Janus','PDB','Categories']

    for source_type in type_list:
        if source_type not in exist_list:
            row = Sourcetype(source_type = source_type)
            session.add(row)

    session.commit()
    #flash('File Type Initialized', category='message')
  
    

    return