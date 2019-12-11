from flask import render_template, redirect, request, flash
from flask_appbuilder.models.sqla.interface import SQLAInterface
from flask_appbuilder import ModelView, BaseView, expose, has_access, action
from app import appbuilder, db
from flask_appbuilder.models.sqla.filters import FilterEqual, FilterNotContains, FilterGreater
from .models import Doc_list, Janus, Pdb, Mscode, Category, SourceFiles, Sourcetype, Janusms, Mdi

from .helpers import (janus_upload, document_list_upload, pdb_list_upload,
                      category_upload, update_all, check_pdb_not_in_janus,
                      init_file_type, update_rq)
from flask_appbuilder.filemanager import get_file_original_name
from config import UPLOAD_FOLDER
from .full_mdi import *
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

'''
class WrongPdbView(ModelView):
    datamodel = SQLAInterface(Pdb)
    list_columns = ['ex_client_reference', 'doc_reference',
                    'revision_number', 'required_action']

    @action("blacklist", "Black List", "All Documents -> Black List (No Progress), Really?", "fa-rocket")
    def blacklist(self, items):
        for item in items:
            doc = BlackList(
                client_reference=item.ex_client_reference,
                doc_reference=item.doc_reference,
                title=item.title,
                source_type=Sourcetype(source_type='PDB')
            )
            
            self.datamodel.add(doc)
        print('Black List action on LIST')
        self.datamodel.delete_all(items)
        self.update_redirect()
        
        return redirect(self.get_redirect())

'''
from .helpers import test_rq, fire_msg

class PdbView(ModelView):
    datamodel = SQLAInterface(Pdb)
    list_columns = ['ex_client_reference', 'client_reference', 'doc_reference',
                    'revision_number', 'document_revision_object']
 
    edit_columns = ['ex_client_reference', 'client_reference', 'doc_reference',
                    'title','revision_number', 'document_revision_object']

    @action("add_mdi", "Add to MDI", "All Documents -> to MDI List, Really?", "fa-rocket")
    def add_mdi(self, items):
        #session = db.session
        janus_count = 0
        pdb_count = 0
        doc_count = 0
        if isinstance(items, list):
            for item in items:
                print(item.client_reference)
                print(item.ex_client_reference)
                cat_code = item.ex_client_reference.split('-')[1][1:]
                print(cat_code)
                category = db.session.query(Category).filter(Category.code == cat_code).first()
                if category:    
                    doc = Doc_list(
                        cat= cat_code,
                        title=item.title,
                        org=item.discipline,
                        cat_class=category.document_class,
                        #class_two=row[7].value,
                        #weight=row[8].value,
                        doc_reference=item.doc_reference
                    )
                    print(item.ex_client_reference)
                    doc.client_reference = item.ex_client_reference
                    self.datamodel.add(doc)
                    doc_count += 1

                    # Update PDB references to new Doc
                    item.client_reference_id = item.ex_client_reference
                    item.ex_client_reference = None

                    # Update Janus REferences
                    janus = db.session.query(Janus).filter(Janus.doc_reference == item.doc_reference).first()
                    
                    if janus:
                        janus.client_reference_id = item.client_reference_id
                        janus.ex_client_reference = None
                        janus_count += 1
                        self.datamodel.edit(janus)
                    
                    self.datamodel.edit(item)
                    
                    pdb_count += 1
                else:
                    flash(str(item.ex_client_reference) + ' Category Code ' + str(cat_code)+ ' Not Found.',category='warning')
            
            flash(str(doc_count) + ' Documents ' + str(pdb_count) + ' PDB revisions and ' + str(janus_count) + ' Janus have been updated.', category='info')
            self.update_redirect()
        else:
            doc = BlackList(
                client_reference=item.client_reference,
                doc_reference=item.doc_reference,
                title=item.title
            )
            self.datamodel.add(doc)

        return redirect(self.get_redirect())
    
    @action("redis", "Test RQ", "Test START, Really?", "fa-rocket")
    def redis(self, items): 
        test_rq(items.title)
        return redirect(self.get_redirect())

    @action("message", "Message RQ", "Test START, Really?", "fa-rocket")
    def message(self, items): 
        fire_msg(self,items.title)
        return redirect(self.get_redirect())

    

class JanusView(ModelView):
    datamodel = SQLAInterface(Janus)
    list_columns = ['doc_reference','mscode','pdb_issue',  'initial_plan_date',
                    'revised_plan_date', 'forecast_date', 'actual_date','planned_date']
    @action("add_mdi", "Add to MDI", "All Documents -> to MDI List, Really?", "fa-rocket")
    def add_mdi(self, items):
        #session = db.session
        janus_count = 0
        pdb_count = 0
        doc_count = 0
        if isinstance(items, list):
            for item in items:
                print(item.client_reference)
                print(item.ex_client_reference)
                cat_code = item.ex_client_reference.split('-')[1][1:]
                print(cat_code)
                category = db.session.query(Category).filter(Category.code == cat_code).first()
                if category:    
                    doc = Doc_list(
                        cat= cat_code,
                        title=item.title,
                        org=item.discipline,
                        cat_class=category.document_class,
                        #class_two=row[7].value,
                        #weight=row[8].value,
                        doc_reference=item.doc_reference
                    )
                    print(item.ex_client_reference)
                    doc.client_reference = item.ex_client_reference
                    self.datamodel.add(doc)
                    doc_count += 1

                    # Update PDB references to new Doc
                    item.client_reference_id = item.ex_client_reference
                    item.ex_client_reference = None

                    # Update Janus REferences
                    janus = db.session.query(Janus).filter(Janus.doc_reference == item.doc_reference).first()
                    
                    if janus:
                        janus.client_reference_id = item.client_reference_id
                        janus.ex_client_reference = None
                        janus_count += 1
                        self.datamodel.edit(janus)
                    
                    self.datamodel.edit(item)
                    
                    pdb_count += 1
                else:
                    flash(str(item.ex_client_reference) + ' Category Code ' + str(cat_code)+ ' Not Found.',category='warning')
            
            flash(str(doc_count) + ' Documents ' + str(pdb_count) + ' PDB revisions and ' + str(janus_count) + ' Janus have been updated.', category='info')
            self.update_redirect()
        else:
            doc = BlackList(
                client_reference=item.client_reference,
                doc_reference=item.doc_reference,
                title=item.title
            )
            self.datamodel.add(doc)

        return redirect(self.get_redirect())

class DocumentListView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    add_columns = ['doc_reference', 'client_reference', 'title']
    list_columns = ['doc_reference', 'client_reference', 'title']
    show_columns = ['doc_reference', 'client_reference', 'title',
                    'cat', 'org', 'weight', 'cat_class', 'class_two','mdi']
    search_columns = ['client_reference','note']
    related_views = [JanusView, PdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'

    @action("followOff", "Do Not Follow this Document", "All selected Documents -> Follow = FALSE, Really?", "fa-rocket")
    def followOff(self, items):
        for item in items:
            item.mdi = False
            
            self.datamodel.edit(item)
        self.update_redirect()
        
        return redirect(self.get_redirect())

class DocumentListView2(ModelView):
    datamodel = SQLAInterface(Doc_list)


class JanusmsView(ModelView):
    datamodel = SQLAInterface(Janusms)


class MscodeView(ModelView):
    datamodel = SQLAInterface(Mscode)
    list_columns = ['position','mscode','description','mdi']
    add_columns = ['position','mscode','description','mdi', 'JanusMS']
    related_views = [JanusmsView]


class CategoryView(ModelView):
    datamodel = SQLAInterface(Category)
    list_columns = ['sheet_name', 'code', 'description']


'''
class WrongDocRefView(ModelView):
    datamodel = SQLAInterface(Doc_list)
    base_filters = [['client_reference', FilterEqual, 'wrong_client_ref']]
    related_views = [JanusView, WrongPdbView]
    show_template = 'appbuilder/general/model/show_cascade.html'
    list_columns = ['title']
    show_columns = ['title']
'''
'''
class BlackListView(ModelView):
    datamodel = SQLAInterface(BlackList)
    list_columns = ['client_reference', 'doc_reference', 'title']
'''

class Setting_updateView(BaseView):
    default_view = 'upload_setting'

    @expose('/setting/', methods=['GET'])
    @has_access
    def upload_setting(self):
        doc, pdb, janus, cat, deleted_doc, deleted_pdb, deleted_janus, deleted_cat = update_all()
        pdb_not_in_janus = check_pdb_not_in_janus()
        return self.render_template('setting_status.html',
                                    doc=doc,
                                    pdb=pdb,
                                    janus=janus,
                                    cat=cat,
                                    deleted_doc=deleted_doc,
                                    deleted_pdb=deleted_pdb,
                                    deleted_janus=deleted_janus,
                                    deleted_cat=deleted_cat,
                                    pdb_not_in_janus=pdb_not_in_janus
                                    )


class SourceFileTypeView(ModelView):
    datamodel = SQLAInterface(Sourcetype)
    list_columns = ['source_type', 'description']

    

class SourceFilesView(ModelView):
    datamodel = SQLAInterface(SourceFiles)
    add_columns = ['source_type', 'file_source', 'description']
    list_columns = ['source_type', 'file_source', 'description']

    @action("Update2", "Update DB2", "Delete All Data and Update by source File, Really?", "fa-rocket", single=True, multiple=False)
    def update2(self, items): 
        update_rq(items.source_type)
        return redirect(self.get_redirect())

class MdiView(ModelView):
    datamodel = SQLAInterface(Mdi)
    list_columns = ['name','description','download']
    add_columns = ['name','description']
    

from .helpers import mdi_rq
class MyView(BaseView):
    route_base = "/mdi"
    
    @expose('/new/')
    def mdi_new(self):
        result = mdi_rq()
        return result

    @expose('/method1/<string:param1>')
    def method1(self, param1):
        # do something with param1 
        # and return it
        return param1

    @expose('/method2/<string:param1>')
    def method2(self, param1):
        # do something with param1
        # and render it
        param1 = 'Hello %s' % (param1)
        return param1

appbuilder.add_view_no_menu(MyView())



@appbuilder.app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html', base_template=appbuilder.base_template, appbuilder=appbuilder), 404


db.create_all()

appbuilder.add_view(DocumentListView, "Document", icon="fa-folder-open-o",
                    category="List", category_icon='fa-envelope')
appbuilder.add_view(JanusView, "Janus", icon="fa-folder-open-o",
                    category="List", category_icon='fa-envelope')
appbuilder.add_view(PdbView, "PDB", icon="fa-folder-open-o",
                    category="List", category_icon='fa-envelope')
appbuilder.add_view(MdiView, "MDI", icon="fa-folder-open-o",
                    category="List", category_icon='fa-envelope')




'''
# appbuilder.add_view(BlackListView, "Black List",)

                    icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
# appbuilder.add_view(WrongDocRefView, "Wrong References",
                    icon="fa-folder-open-o", category="List", category_icon='fa-envelope')
# appbuilder.add_view_no_menu(WrongPdbView)
'''

appbuilder.add_view(SourceFilesView, "Source File", icon="fa-folder-open-o",
                    category="Setting", category_icon='fa-envelope')
appbuilder.add_view(SourceFileTypeView, "File Type", icon="fa-folder-open-o",
                    category="Setting", category_icon='fa-envelope')
appbuilder.add_view(MscodeView, "Milestones", icon="fa-folder-open-o",
                    category="Setting", category_icon='fa-envelope')
appbuilder.add_view(JanusmsView, "Janus MS", icon="fa-folder-open-o",
                    category="Setting", category_icon='fa-envelope')

appbuilder.add_view(CategoryView, "Category", icon="fa-folder-open-o",
                    category="Setting", category_icon='fa-envelope')


appbuilder.add_separator(category='Setting')


appbuilder.add_view(Setting_updateView, "Setting Update",
                    icon="fa-folder-open-o", category="Setting", category_icon='fa-envelope')

appbuilder.add_link('Document Code Error', '/documentlistview/list/?_flt_0_note=Document+Code+Error',
                    icon="fa-folder-open-o", category="DCC Check", category_icon='fa-envelope')
appbuilder.add_link('PDB Not in Document List', '/pdbview/list/?_flt_0_client_reference=__None',
                    icon="fa-folder-open-o", category="DCC Check", category_icon='fa-envelope')
appbuilder.add_link('Janus Not in Document List', '/janusview/list/?_flt_0_client_reference=__None',
                    icon="fa-folder-open-o", category="DCC Check", category_icon='fa-envelope')
# document_list_upload()
# janus_upload()
# pdb_list_upload()
# category_upload()
# update_all()
# check_pdb_not_in_janus()
# init_file_type()
