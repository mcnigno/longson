from .models import Doc_list, Pdb, Janus, Mscode
from app import db



def full_mdi(client_reference): 
    session = db.session

    document = session.query(Doc_list).filter(Doc_list.client_reference == client_reference).first()
    #print('Before PDB Query for:', client_reference, 'len',len(client_reference) )
    pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == client_reference ).order_by(Pdb.revision_number).all()
    
    janus_document = session.query(Janus).filter(Janus.client_reference_id == client_reference).all()
    mscode_list = session.query(Mscode).order_by(Mscode.position).all()
    #mscode_list = [x.code for x in mscodes]
    ### Order PDB revision by letters then number | fake solution... Z0, Z1 etc..
    ### Crea un dict con la posizione per ogni mscode
    ### Aggiungi la posizione alla lista ordinata
    mscode_dict = dict([(x.mscode, x.position) for x in mscode_list])
    
    #print(mscode_dict)
    ordered_list = []
    pdb_revision_object_list = set()  
    if pdb_document:
        for doc in pdb_document:
            #print('full doc:', doc.client_reference,doc.revision_number,doc.document_revision_object)
            if doc is not None \
                and doc.client_reference \
                and doc.revision_number \
                and doc.document_revision_object:
                #print('issue, revision :', doc.document_revision_object + doc.revision_number)
                try:
                    rev = int(doc.revision_number)
                    mscode = doc.document_revision_object
                    ordered_list.append((
                        str(mscode_dict[doc.document_revision_object]) +
                        doc.document_revision_object+ 'Z'+ str(rev),
                        doc  ))
                except:
                    ordered_list.append((
                        str(mscode_dict[doc.document_revision_object]) +
                        doc.document_revision_object + 
                        doc.revision_number,doc )) 
                
                pdb_revision_object_list.add(doc.document_revision_object)
    else:
        print(' <<<<<<<<<<<<< NOT FOUND IN PDB   |||||| <<<<<<<<<<<<<')
    #print(sorted(ordered_list))
    if janus_document:
        print(' ------------- FOUND IN JANUS', client_reference, ' |-|-|-|-|-| <<<<<<<<<<<<<')
        for j in janus_document:
            print('CLIENT REF, PDB ISUUE:',client_reference, j.pdb_issue)
            print(pdb_revision_object_list)
            #print(mscode_list)
            if j.mscode != 'START' \
                and j.pdb_issue != 'NOT APPLIC' \
                and j.pdb_issue not in pdb_revision_object_list \
                and client_reference \
                and j.pdb_issue in mscode_dict:
                
                print('JANUS OK')
                #and doc.client_reference_id:
                #print('client reference',doc.client_reference_id)
                #print('vvvvvvvv')
                fake_doc = Pdb(client_reference_id=client_reference)
                fake_doc.document_revision_object = j.pdb_issue
        
                #print('Before ORDER ---------',j.pdb_issue)
                #print(mscode_dict)
                #print(ordered_list)
                print('Janus Issue Order', str(mscode_dict[j.pdb_issue]) + j.mscode)
                ordered_list.append((
                    str(mscode_dict[j.pdb_issue]) +
                    j.mscode,
                    fake_doc))
                #print(j.mscode)
                '''
                print('After ORDER ---------',fake_doc.client_reference_id,
                                fake_doc.document_revision_object,
                                j.pdb_issue,
                                type(j.pdb_issue)
                                )
                #print(ordered_list)
                '''
                ''' 
                for x,doc in sorted(ordered_list):
                    print(x)
                '''
            else: 
                print('ERROR in JANUS')
    else:
        print(' >>>>>< > > > <  NOT FOUND IN JANUS  ** J ** <<<<<<<<<<<<<')
    return sorted(ordered_list)  

#full_mdi() 
def full_mdi_last_rev(client_reference): 
    session = db.session

    document = session.query(Doc_list).filter(Doc_list.client_reference == client_reference).first()

    pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == client_reference ).order_by(Pdb.revision_number).all()
    janus_document = session.query(Janus).filter(Janus.client_reference_id == client_reference).all()
    mscode_list = session.query(Mscode).order_by(Mscode.position).all()
    ### Order PDB revision by letters then number | fake solution... Z0, Z1 etc..
    ### Crea un dict con la posizione per ogni mscode
    ### Aggiungi la posizione alla lista ordinata
    mscode_dict = dict([(x.mscode, str(x.position).zfill(2)) for x in mscode_list])
    mscode_actual = set([x.document_revision_object for x in pdb_document]+[x.mscode for x in janus_document])
    
    mscode_actual2 = set([x.document_revision_object for x in pdb_document])

    ordered_list = []
    pdb_revision_object_list = set()
    
    relations = {
        'IFD': ['IFD1','IFD2'],
        'IFA': ['IFR']
    }

    for code in mscode_actual:
        pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == client_reference,
                                                Pdb.document_revision_object == code ).order_by(Pdb.revision_number).all()
        if pdb_document:
            last_doc = pdb_document[-1]
            pdb_document[-1].note = 'last'
            print('LAST',last_doc.client_reference_id,last_doc.revision_number,last_doc.document_revision_object)
    #print(mscode_dict)
      
        if pdb_document:
            for doc in pdb_document:
                print('full doc:', doc.client_reference,doc.revision_number,doc.document_revision_object)
                if doc is not None \
                    and doc.client_reference \
                    and doc.revision_number \
                    and doc.document_revision_object:
                    try:
                        rev = int(doc.revision_number)
                        mscode = doc.document_revision_object
                        ordered_list.append((
                            'Z'+ str(rev) +
                            
                            str(mscode_dict[doc.document_revision_object]) +
                            doc.document_revision_object,
                            doc  ))
                    except:
                        ordered_list.append((
                            doc.revision_number +
                            
                            str(mscode_dict[doc.document_revision_object]) +
                            doc.document_revision_object,
                            doc )) 
                    
                    pdb_revision_object_list.add(doc.document_revision_object)
                    print(doc.note)
        #print(sorted(ordered_list))
    
    for j in janus_document:
        if j.mdi == True \
        and j.mscode not in pdb_revision_object_list:
        #and doc.client_reference_id:
            #print('client reference',doc.client_reference_id)
            #print('vvvvvvvv')
            fake_doc = Pdb(client_reference_id=client_reference)
            fake_doc.document_revision_object = j.mscode
            fake_doc.note = 'last'
    
            
            ordered_list.append((
                #'Z' +
                str(mscode_dict[j.mscode]) +
                j.mscode,
                fake_doc))
            #print(j.mscode)
    print('XXXXXXX   ORDERED LIST XXXXXXXX')
    for x,doc in sorted(ordered_list):
        print(x)
    return sorted(ordered_list) 



def fulmdi2(client_reference='OL1-2M03-2049'):

    session = db.session
    document = session.query(Doc_list).filter(Doc_list.client_reference == client_reference).first()
    #print('Before PDB Query for:', client_reference, 'len',len(client_reference) )
    pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == client_reference ).order_by(Pdb.revision_number).all()
    
    janus_document = session.query(Janus).filter(Janus.client_reference_id == client_reference).all()
    #mscode_list = session.query(Mscode).order_by(Mscode.position).all()
    
    #mscode_list = [x.code for x in mscodes]
    ### Order PDB revision by letters then number | fake solution... Z0, Z1 etc..
    ### Crea un dict con la posizione per ogni mscode
    ### Aggiungi la posizione alla lista ordinata
    #mscode_dict = dict([(x.mscode, x.position) for x in mscode_list])
    
    #print(mscode_dict)
    ordered_list = []
    pdb_revision_object_list = set() 

    
    pdb_document = session.query(Pdb).filter(
            Pdb.client_reference_id == client_reference 
            ).order_by(Pdb.transmittal_date).all()
    
    if pdb_document:
        for doc in pdb_document:
            if doc.transmittal_date == None: doc.transmittal_date = doc.actual_response_date
            ordered_list.append((doc.transmittal_date, doc))
            pdb_revision_object_list.add(doc.document_revision_object)
    
    janus_document = session.query(Janus).filter(
            Janus.client_reference_id == client_reference
            ).all()
    
    if janus_document:
        for doc in janus_document:
            if doc.pdb_issue not in pdb_revision_object_list \
                and doc.pdb_issue !='NOT APPLIC':
                fake_doc = Pdb(client_reference_id=client_reference)
                fake_doc.document_revision_object = doc.pdb_issue
                ordered_list.append((doc.planned_date, fake_doc))
    
    
    return sorted(ordered_list, key=lambda date: date[0])
    #return ordered_list.sort(key=operator.itemgetter(1))

#fulmdi2()