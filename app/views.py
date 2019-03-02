from flask import render_template
from flask_appbuilder.models.sqla.interface import SQLAInterface
from flask_appbuilder import ModelView
from app import appbuilder, db
from flask_appbuilder.models.sqla.filters import FilterEqual
from .models import Doc_list, Janus, Pdb, Mscode, Category

from .helpers import janus_upload, document_list_upload, pdb_list_upload, category_upload

"""
    Create your Views::


    class MyModelView(ModelView):
        datamodel = SQLAInterface(MyModel)


    Next, register your Views::


    appbuilder.add_view(MyModelView, "My View", icon="fa-folder-open-o", category="My Category", category_icon='fa-envelope')
"""

"""
    Application wide 404 error handler
"""
class PdbView(ModelView):
    datamodel = SQLAInterface(Pdb)
    list_columns = ['client_reference','revision_number', 'required_action']

class JanusView(ModelView):
    datamodel = SQLAInterface(Janus)
    list_columns = ['client_reference','title','mscode','initial_plan_date',
                    'revised_plan_date','forecast_date','actual_date']

class DocumentListView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    add_columns = ['doc_reference','client_reference','title']
    list_columns = ['doc_reference','client_reference','title']
    show_columns = ['doc_reference','client_reference','title',
                    'cat','org','weight','cat_class','class_two']
    related_views = [JanusView, PdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'

class MscodeView(ModelView):
    datamodel = SQLAInterface(Mscode)

class CategoryView(ModelView):
    datamodel = SQLAInterface(Category)
    list_columns = ['sheet_name','code','description']

class WrongReferencesView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    base_filters = [['doc_reference', FilterEqual, 'wrong_ref']]
    related_views = [JanusView, PdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'
    list_columns = ['title']
    show_columns = ['title']

@appbuilder.app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html', base_template=appbuilder.base_template, appbuilder=appbuilder), 404

db.create_all()
appbuilder.add_view(DocumentListView, "Document", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(JanusView, "Janus", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(PdbView, "PDB", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(MscodeView, "Milestones", icon="fa-folder-open-o", category="Setting", category_icon='fa-envelope')
appbuilder.add_view(WrongReferencesView, "Wrong References", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(CategoryView, "Category", icon="fa-folder-open-o", category="Setting", category_icon='fa-envelope')


#document_list_upload()
#janus_upload()
#pdb_list_upload()
#category_upload()

 
 
