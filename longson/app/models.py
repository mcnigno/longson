from flask_appbuilder import Model
from flask_appbuilder.models.mixins import AuditMixin, FileColumn, ImageColumn
from sqlalchemy import Column, Integer, String, ForeignKey, Date 
from sqlalchemy.orm import relationship
"""

You can use the extra Flask-AppBuilder fields and Mixin's

AuditMixin will add automatic timestamp of created and modified by who

TITLE	MILESTONE CHAIN	DBS	WBS	FREE CRITERIA 1	FREE CRITERIA 2	FREE CRITERIA 3	FREE CRITERIA 4	MSCODE	CUMULATIVE	MANUFACTURING CUMULATIVE	EMPTY COL	OBS	INITIAL PLAN DATE	REVISED PLAN DATE	FORECAST DATE	ACTUAL DATE
"""
        
class Janus(Model):
    id = Column(Integer, primary=True)
    linenumber = Column(Integer)
    cat = Column(String(4))
    doc_reference = Column(String(50))
    client_reference = Column(String(50))
    weight = Column(Integer)
    title = Column(String)
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








