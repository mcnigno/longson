from .models import Doc_list, Pdb, Janus, Mscode
from app import db



def full_mdi(client_reference): 
    session = db.session

    document = session.query(Doc_list).filter(Doc_list.client_reference == client_reference).first()

    pdb_document = session.query(Pdb).filter(Pdb.client_reference_id == client_reference ).order_by(Pdb.revision_number).all()
    janus_document = session.query(Janus).filter(Janus.client_reference_id == client_reference).all()
    mscode_list = session.query(Mscode).order_by(Mscode.position).all()
    ### Order PDB revision by letters then number | fake solution... Z0, Z1 etc..
    ### Crea un dict con la posizione per ogni mscode
    ### Aggiungi la posizione alla lista ordinata
    mscode_dict = dict([(x.mscode, x.position) for x in mscode_list])

    #print(mscode_dict)
    ordered_list = []
    pdb_revision_object_list = set()  
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
                        str(mscode_dict[doc.document_revision_object]) +
                        doc.document_revision_object+ 'Z'+ str(rev),
                        doc  ))
                except:
                    ordered_list.append((
                        str(mscode_dict[doc.document_revision_object]) +
                        doc.document_revision_object + 
                        doc.revision_number,doc )) 
                
                pdb_revision_object_list.add(doc.document_revision_object)
    #print(sorted(ordered_list))
    
    for j in janus_document:
        if j.mscode != 'START' \
        and j.mscode not in pdb_revision_object_list:
        #and doc.client_reference_id:
            #print('client reference',doc.client_reference_id)
            #print('vvvvvvvv')
            fake_doc = Pdb(client_reference_id=client_reference)
            fake_doc.document_revision_object = j.mscode
    
            
            ordered_list.append((
                str(mscode_dict[j.mscode]) +
                j.mscode,
                fake_doc))
            #print(j.mscode)
    
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