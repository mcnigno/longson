from flask_appbuilder import Model
from flask_appbuilder.models.mixins import AuditMixin, FileColumn, ImageColumn
from sqlalchemy import Column, Integer, String, ForeignKey, Date, Text, Boolean
from sqlalchemy.orm import relationship



"""

You can use the extra Flask-AppBuilder fields and Mixin's

AuditMixin will add automatic timestamp of created and modified by who

"""


class Doc_list(Model):
    #id = Column(Integer, primary_key=True, autoincrement=True)
    cat = Column(String(4))
    client_reference = Column(String(50), primary_key=True, unique=True)
    doc_reference = Column(String(50))
    title = Column(String(255))
   
    org = Column(String(50))
    cat_class = Column(String(50))
    class_two = Column(String(50))
    weight = Column(Integer)
    mdi = Column(Boolean, default=True)
    note = Column(String(255))

    project = Column(String(50))
    unit = Column(String(50))
    doc_type = Column(String(50))
    m_class = Column(String(50))
    prog = Column(String(50))

    
    def __init__(self, **kwargs):
        super(Doc_list, self).__init__(**kwargs)
        doc_ref_fields = str(self.doc_reference).split('-')
        if len(doc_ref_fields) == 5:
            self.project = doc_ref_fields[0]
            self.unit = doc_ref_fields[1]
            self.doc_type = doc_ref_fields[2]
            self.m_class = doc_ref_fields[3]
            self.prog = doc_ref_fields[4]
        else:
            self.note = 'Document Code Error' 
            self.mdi = False

    
    def __repr__(self):
        return self.client_reference
    
    
    '''
    def pdb_items(self):
        return int(len(self.pdb))
    
    def janus_items(self):
        return len(self.janus)
    '''



class Pdb(Model):
    id = Column(Integer, primary_key=True)
    client_reference_id = Column(String(50), ForeignKey('doc_list.client_reference'))
    client_reference = relationship(Doc_list)#, backref='pdb', lazy='selectin')
    ex_client_reference = Column(String(50))
    title = Column(String(255))
    revision_number = Column(String(10))
    revision_date = Column(Date)
    document_revision_object = Column(String(255))
    doc_reference = Column(String(255))
    discipline = Column(String(5))
    transmittal_date = Column(Date)
    transmittal_reference = Column(String(255))
    specific_transmittal_number = Column(String(255))
    required_action = Column(String(10))
    response_due_date = Column(Date)
    actual_response_date = Column(Date)
    document_status = Column(String(255))
    client_transmittal_ref_number = Column(String(255))
    remarks = Column(String(255))
    note = Column(String(255))
    document_class = Column(String(10))

    def __init__(self, **kwargs):
        super(Pdb, self).__init__(**kwargs)
        try:
            rev = int(self.revision_number)
            self.revision_number = 'Z' + str(rev)
        except:

            self.revision_number = self.revision_number

    def __repr__(self):
        return str(self.client_reference) + ' Rev. ' + self.revision_number


class Janus(Model):
    id = Column(Integer, primary_key=True)
    linenumber = Column(String(4))
    cat = Column(String(4))

    client_reference_id = Column(String(50), ForeignKey('doc_list.client_reference'))
    client_reference = relationship(Doc_list)#, backref='janus', lazy='selectin')
    ex_client_reference = Column(String(50))

    doc_reference = Column(String(50))
    weight = Column(Integer)
    title = Column(String(255))
    milestone_chain = Column(String(50))
    dbs = Column(String(50))
    wbs = Column(String(50))
    fcr_one = Column(String(50))
    fcr_two = Column(String(50))
    fcr_three = Column(String(50))
    fcr_four = Column(String(50))
    mscode = Column(String(50))
    cumulative = Column(String(50))
    obs = Column(String(50))

    initial_plan_date = Column(Date)
    revised_plan_date = Column(Date)
    forecast_date = Column(Date)
    actual_date = Column(Date)
    planned_date = Column(Date)

    note = Column(String(255))

    pdb_issue = Column(String(10))

    pdb_id = Column(Integer, ForeignKey('pdb.id'))
    pdb = relationship(Pdb)

    def __repr__(self):
        return str(self.client_reference) + self.mscode

 

class Mscode(Model):
    id = Column(Integer, primary_key=True)
    position = Column(Integer)
    mscode = Column(String(10))
    description = Column(String(255))
    mdi = Column(Boolean)
    

class Janusms(Model):
    id = Column(Integer, primary_key=True)
    position = Column(Integer)
    mscode = Column(String(10))
    description = Column(String(255))
    mdi = Column(Boolean)

    mscode_id = Column(Integer, ForeignKey('mscode.id'))
    mscode = relationship(Mscode, backref='JanusMS')

    

class Category(Model):
    id = Column(Integer, primary_key=True)
    sheet_name = Column(String(10))
    code = Column(String(255))
    information = Column(String(255))
    description = Column(Text)
    document_class = Column(String(10))


class Sourcetype(Model):
    id = Column(Integer, primary_key=True)
    source_type = Column(String(50))
    description = Column(Text)

    def __repr__(self):
        return self.source_type


class SourceFiles(Model, AuditMixin):
    id = Column(Integer, primary_key=True)
    file_source = Column(FileColumn())
    source_type_id = Column(Integer, ForeignKey('sourcetype.id'),nullable=False)
    source_type = relationship(Sourcetype)
    description = Column(Text)

    def __repr__(self):
        return self.file_source

from flask import url_for, Markup
class Mdi(Model,AuditMixin):
    id = Column(Integer, primary_key=True)
    file = Column(FileColumn())
    name = Column(String(50),nullable=False)
    description = Column(Text)
    def __repr__(self):
        return self.name
    
    def download(self):
        return Markup(
            '<a href="'
            + url_for("MdiView.download", filename=str(self.file))
            + '">Download</a>'
        )

'''
class BlackList(Model):
    id = Column(Integer, primary_key=True)
    doc_reference = Column(String(50))
    client_reference = Column(String(50))
    title = Column(String(255))
    source_id = Column(Integer, ForeignKey('sourcetype.id'), nullable=False) 
    source_type = relationship(Sourcetype)
    
    def __repr__(self):
        return self.client_reference

'''










