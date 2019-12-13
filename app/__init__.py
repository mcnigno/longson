import logging
from flask import Flask
from flask_appbuilder import SQLA, AppBuilder
from flask_rq2 import RQ
from app.index import MyIndexView


"""
 Logging configuration
"""

logging.basicConfig(format='%(asctime)s:%(levelname)s:%(name)s:%(message)s')
logging.getLogger().setLevel(logging.DEBUG)

app = Flask(__name__)
app.config.from_object('config')
db = SQLA(app)

rq = RQ(app)
#rq.get_scheduler(interval=10)
#rq.get_worker('default','low')

appbuilder = AppBuilder(app, db.session, 
                    #base_template='mybase.html',
                    #indexview=MyIndexView
                    )


"""
from sqlalchemy.engine import Engine
from sqlalchemy import event

#Only include this for SQLLite constraints
@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    # Will force sqllite contraint foreign keys
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()
"""    

from app import views

