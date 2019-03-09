from flask import render_template, redirect, request
from flask_appbuilder.models.sqla.interface import SQLAInterface
from flask_appbuilder import ModelView, BaseView, expose, has_access, action
from app import appbuilder, db
from flask_appbuilder.models.sqla.filters import FilterEqual, FilterNotContains, FilterGreater
from .models import Doc_list, Janus, Pdb, Mscode, Category, PDB_no_Progress

from .helpers import (janus_upload, document_list_upload, pdb_list_upload,
                    category_upload, update_all, check_pdb_not_in_janus)

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
    list_columns = ['client_reference', 'doc_reference','revision_number', 'required_action']
    @action("noProgress", "Not in Progress", "All Documents -> Not in Progress List, Really?", "fa-rocket")
    def noProgress(self, items):
        #session = db.session
        if isinstance(items, list):
            for item in items:
                doc = PDB_no_Progress(
                    client_reference = item.client_reference,
                    doc_reference = item.doc_reference,
                    title = item.title
                )
                self.datamodel.add(doc)
            

            self.update_redirect()
        else:
            doc = PDB_no_Progress(
                    client_reference = item.client_reference,
                    doc_reference = item.doc_reference,
                    title = item.title
                )
            self.datamodel.add(doc)
            
        return redirect(self.get_redirect())

class JanusView(ModelView):
    datamodel = SQLAInterface(Janus)
    list_columns = ['client_reference','doc_reference','mscode','initial_plan_date',
                    'revised_plan_date','forecast_date','actual_date']

class DocumentListView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    add_columns = ['doc_reference','client_reference','title']
    list_columns = ['doc_reference','client_reference','title']
    show_columns = ['doc_reference','client_reference','title',
                    'cat','org','weight','cat_class','class_two']
    search_columns = ['client_reference']
    related_views = [JanusView, PdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'

class DocumentListView2(ModelView):
    datamodel = SQLAInterface(Doc_list)

class MscodeView(ModelView):
    datamodel = SQLAInterface(Mscode)

class CategoryView(ModelView):
    datamodel = SQLAInterface(Category)
    list_columns = ['sheet_name','code','description']

class WrongDocRefView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    base_filters = [['client_reference', FilterEqual, 'wrong_client_ref']]
    related_views = [JanusView, PdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'
    list_columns = ['title']
    show_columns = ['title']



class PDB_no_ProgressView(ModelView):
    datamodel = SQLAInterface(PDB_no_Progress)
    list_columns = ['client_reference','doc_reference', 'title']

class Setting_updateView(BaseView):
    default_view = 'upload_setting'

    @expose('/setting/', methods=['POST','GET'])
    @has_access
    def upload_setting(self):
        if request.method == 'POST':
            doc, pdb, janus, cat = update_all()
            return self.render_template('setting_up.html',
                                        filename=filename_list,
                                        updated_list=updated_list,
                                        count_updated=len(updated_list),
                                        reserved_list=reserved_list,
                                        count_reserved=len(reserved_list))
          
        return self.render_template('setting_up.html')



@appbuilder.app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html', base_template=appbuilder.base_template, appbuilder=appbuilder), 404

db.create_all()
appbuilder.add_view(DocumentListView, "Document", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')

appbuilder.add_view(JanusView, "Janus", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(PdbView, "PDB", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(PDB_no_ProgressView, "PDB Not In Progress", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')

appbuilder.add_view(MscodeView, "Milestones", icon="fa-folder-open-o", category="Setting", category_icon='fa-envelope')
appbuilder.add_view(WrongDocRefView, "Wrong References", icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
appbuilder.add_view(CategoryView, "Category", icon="fa-folder-open-o", category="Setting", category_icon='fa-envelope')
 

#document_list_upload()
#janus_upload()
#pdb_list_upload()
#category_upload()
#update_all()  
#check_pdb_not_in_janus()  
 
 