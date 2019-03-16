from flask_appbuilder import Model
from flask_appbuilder.models.mixins import AuditMixin, FileColumn, ImageColumn
from sqlalchemy import Column, Integer, String, ForeignKey, Date, Text, Boolean
from sqlalchemy.orm import relationship



"""

You can use the extra Flask-AppBuilder fields and Mixin's

AuditMixin will add automatic timestamp of created and modified by who

TITLE	MILESTONE CHAIN	DBS	WBS	FREE CRITERIA 1	FREE CRITERIA 2	FREE CRITERIA 3	FREE CRITERIA 4	MSCODE	CUMULATIVE	MANUFACTURING CUMULATIVE	EMPTY COL	OBS	INITIAL PLAN DATE	REVISED PLAN DATE	FORECAST DATE	ACTUAL DATE
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

    def __repr__(self):
        return self.client_reference
    '''
    def pdb_items(self):
        return int(len(self.pdb))
    
    def janus_items(self):
        return len(self.janus)
    '''
class Janus(Model):
    id = Column(Integer, primary_key=True)
    linenumber = Column(String(4))
    cat = Column(String(4))

    client_reference_id = Column(String(50), ForeignKey('doc_list.client_reference'), nullable=False)
    client_reference = relationship(Doc_list)#, backref='janus', lazy='selectin')

    doc_reference = Column(String(50))
    weight = Column(Integer)
    title = Column(String(255))
    milestone_chain = Column(String(50))
    dbs = Column(String(50))
    wbs = Column(String(50))
    fcr_one = Column(String(50))
    fcr_two = Column(String(50))
    fcr_three = Column(String(50))
    mscode = Column(String(50))
    cumulative = Column(String(50))
    obs = Column(String(50))
    initial_plan_date = Column(Date)
    revised_plan_date = Column(Date)
    forecast_date = Column(Date)
    actual_date = Column(Date)

    def __repr__(self):
        return str(self.client_reference) + self.mscode


class Pdb(Model):
    id = Column(Integer, primary_key=True)
    client_reference_id = Column(String(50), ForeignKey('doc_list.client_reference'), nullable=False)
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

    def __repr__(self):
        return self.client_reference + ' Rev. ' + self.revision_number

class PDB_no_Progress(Model):
    doc_reference = Column(String(50), primary_key=True, unique=True)
    client_reference = Column(String(50))
    title = Column(String(255))

    def __repr__(self):
        return self.client_reference

class Mscode(Model):
    id = Column(Integer, primary_key=True)
    position = Column(Integer)
    mscode = Column(String(10))
    description = Column(String(255))


class Category(Model):
    id = Column(Integer, primary_key=True)
    sheet_name = Column(String(10))
    code = Column(String(255))
    information = Column(String(255))
    description = Column(Text)

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
    














