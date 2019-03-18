from .models import Janus, Doc_list, Pdb, Category, SourceFiles, Sourcetype, Mscode
from app import db
from datetime import datetime
import openpyxl
from flask import flash
from flask_appbuilder.models.sqla.interface import SQLAInterface
from config import UPLOAD_FOLDER

janus_not_in_document_list = []
pdb_not_in_document_list = []


def janus_upload_from_txt(source):
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
        
                linenumber=row[0],
                cat=row[1],
                doc_reference=row[2],
                weight=row[9],
                title=row[11],
                milestone_chain=row[12],
                dbs=row[13],
                wbs=row[14],
                fcr_one=row[15],
                fcr_two=row[16],
                fcr_three=row[17],
                mscode=row[19],
                cumulative=row[20],
                obs=row[23],
                initial_plan_date=date_parse(row[24]),
                revised_plan_date=date_parse(row[25]),
                forecast_date=date_parse(row[26]),
                actual_date=date_parse(row[27])
            )
        
            janus.created_by_fk = '1'
            janus.client_reference_id = row[8]

            # Check if the janus doc is or not in document list
            
            if row[8] in doc_list:
                session.add(janus)
                count_janus += 1
            else:
                janus.ex_client_reference_id = row[8]
                janus.note = janus.client_reference_id + ' | No reference in Document List.'
                janus.client_reference_id = None
                session.add(janus)
                janus_not_in_document_list.append(janus)

        session.commit()
        return str(count_janus) + ' Janus Updated'
    except:
        return 'Janus FAIL: check your source file'


def janus_upload(source):
    session = db.session
    janus_file = openpyxl.load_workbook(UPLOAD_FOLDER + source, data_only=True, guess_types=True)
    janus_sheet = janus_file.active
    #row_number = 0
    count_janus = 0
    doc_list = [x.client_reference for x in session.query(Doc_list).all()]
    try:
        for row in janus_sheet.iter_rows(min_row=2):
            
            janus = Janus(
        
                linenumber=row[0].value,
                cat=row[1].value,
                doc_reference=row[2].value,
                weight=row[9].value,
                title=row[11].value,
                milestone_chain=row[12].value,
                dbs=row[13].value,
                wbs=row[14].value,
                fcr_one=row[15].value,
                fcr_two=row[16].value,
                fcr_three=row[17].value,
                fcr_four=row[18].value,
                mscode=row[19].value,
                cumulative=row[20].value,
                obs=row[23].value,
                initial_plan_date=row[24].value, 
                revised_plan_date=row[25].value,
                forecast_date=row[26].value,
                actual_date=row[27].value,
                planned_date=row[32].value
            )  
             
            janus.created_by_fk = '1'
            
            janus.client_reference_id = row[8].value
            
            
            # Check if the janus doc is or not in document list 
            
            if row[8].value in doc_list:
                
                session.add(janus)
                count_janus += 1
            else:
                
                janus.ex_client_reference = row[8].value
                janus.note = str(janus.client_reference_id) + ' | No reference in Document List.'
                janus.client_reference_id = None
                
                session.add(janus)
                #janus_not_in_document_list.append(janus)

        session.commit()
        return str(count_janus) + ' Janus Updated'
    except:
        return 'Janus FAIL: check your source file'


def janus_update():
    session = db.session
    sf = SourceFiles
    files = dict([(str(x.source_type), x.file_source) for x in session.query(sf).all()])
    janusdel = session.query(Janus).delete()
    print('deleted from JANUS', janusdel)
    janus = janus_upload(files['Janus'])

    return janus


# janus_test()  
def date_parse(date):
    try:
        if isinstance(date, datetime): return date
        if date == '#N/A': return None
        if date == '' or date is None: return None

        #return datetime.strptime(date, '%d-%b-%y')
    except:
        return None


def document_list_upload(source):
    print('document_list_upload***')
    session = db.session

    doclist = openpyxl.load_workbook(UPLOAD_FOLDER + source)
    doclist_ws = doclist.active
    # count_id = 0
    count_doc = 0
    # wrong_ref = session.query(Doc_list).filter(Doc_list.client_reference == 'wrong_client_ref').first()
    # empty_client_references = session.query(Doc_list).filter(Doc_list.client_reference == 'EmptyClientReferences').first()            
    try:    
        for row in doclist_ws.iter_rows(min_row=2):
            if row[2].value:
                doclist_row = Doc_list(
                    cat=row[0].value,
                    title=row[3].value,
                    org=row[5].value,
                    cat_class=row[6].value,
                    class_two=row[7].value,
                    weight=row[8].value,
                    doc_reference=row[12].value
                )
                doclist_row.created_by_fk = '1'
                doclist_row.client_reference = row[2].value
                count_doc += 1
                session.add(doclist_row)
            
        print('Document in DC processed', count_doc)
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
                doc_reference=row[0].value,
                title=row[1].value,
                revision_number=row[2].value,
                revision_date=date_parse(row[3].value),
                document_revision_object=row[4].value,

                discipline=row[6].value,
                transmittal_date=date_parse(row[7].value),
                transmittal_reference=row[8].value,
                specific_transmittal_number=row[9].value,
                required_action=row[10].value,
                response_due_date=date_parse(row[11].value),
                actual_response_date=date_parse(row[12].value),
                document_status=row[13].value,
                client_transmittal_ref_number=row[14].value,
                remarks=row[15].value,
            )
            pdb.created_by_fk = '1'
            pdb.client_reference_id = row[5].value
            

            
            # Check if the pdb doc is or not in document list
            
            if row[5].value in doc_list:
                session.add(pdb)
                count_pdb += 1
            else:
                pdb.ex_client_reference = pdb.client_reference_id
                pdb.note = pdb.client_reference_id + ' | No reference in Document List.'
                pdb.client_reference_id = None
                
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
    
def mscode_upload(source):
    wcm = openpyxl.load_workbook(UPLOAD_FOLDER + source, data_only=True) 
    wcm = wcm.active
    
    session = db.session
    count_mscodes = 0
    try:
        for row in wcm.iter_rows(min_row=2):
            
            if row[0].value is not None:
                
                mscode = Mscode(

                    mscode=row[0].value,
                    position=int(row[1].value),
                    description=row[2].value,

                )
                session.add(mscode)
                count_mscodes += 1
    
        session.commit()
        return str(count_mscodes) + ' MS Codes updated!'
    except:
        return 'MS Codes FAIL: check your source file.'

def mscode_update():
    session = db.session
    sf = SourceFiles
    files = dict([(str(x.source_type), x.file_source) for x in session.query(sf).all()])
    mscodedel = session.query(Mscode).delete()
    print('deleted from MSCODES', mscodedel)
    
    mscode = mscode_upload(files['MSCodes'])
    print(mscode)

    return mscode

#mscode_update()

def update_all():
    init_file_type()
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
    '''Initialize the DB with all the File Source Type Needed'''
    session = db.session
    exist_list = [x.source_type for x in db.session.query(Sourcetype).all()]
    type_list = ['Document List','Janus','PDB','Categories','MSCodes']

    for source_type in type_list:
        if source_type not in exist_list:
            row = Sourcetype(source_type = source_type)
            session.add(row)

    session.commit()
    #flash('File Type Initialized', category='message')
  
    

    return

###
### MDI Main Document Index to Excel
def mdi_excel():
    session = db.session
    
    # For every Category on each document
    categorie_list = session.query(Category).filter(Category.code != None, Category.information != None).all()
    mscodes_list = session.query(Mscode).order_by(Mscode.position).all()
    #print(mscodes_list)
    for category in categorie_list:
        cat_code = 'Category Code: ' + category.code + ' ' + category.information
        document_list = session.query(Doc_list).filter(Doc_list.cat == category.code).all()
        

        if document_list:      
            #print(cat_code)
            for document in document_list: 
                #print(document)
                pdb_document_list = session.query(Pdb).filter(Pdb.client_reference_id == document.client_reference).all()
                #janus_list = session.query(Janus).filter(Janus.client_reference_id == document.client_reference).all()
                mscode_actions = set([x.required_action for x in pdb_document_list])
                if mscodes_list:
                    for ms in mscode_actions:
                        pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == document.client_reference, Pdb.required_action == ms ).first()
                        #janus_document = session.query(Janus).filter(Janus.client_reference_id == document.client_reference, Janus.mscode == ms.mscode ).first()
                        
                        if pdb_document:
                            print('PDB:',pdb_document.client_reference,pdb_document.required_action)
                        #if janus_document:
                        #    print('JANUS:',janus_document.client_reference,janus_document.mscode)
                ''' 
                for milestone in set([x.required_action for x in pdb_document_list] + [x.mscode for x in janus_list]):
                    if milestone == code.mscode:    
                        print('pdb + janus milestones:', document, milestone)    
                        if pdb_document_list:
                            for doc in pdb_document_list:
                                print(cat_code, document, doc.title)
                                if janus_list:
                                    
                                    for milestone in janus_list[1:]:
                                        print(milestone.mscode) 
                '''                        
    return  
#mdi_excel() 
# To Do: Check if PDB has 
