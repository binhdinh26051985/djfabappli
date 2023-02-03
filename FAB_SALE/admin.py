from django.contrib import admin
import import_export
from .models import *
from import_export.widgets import ForeignKeyWidget,ManyToManyWidget,DateWidget
from import_export import fields, resources
from import_export.admin import ImportExportModelAdmin

class PODetailsAdminResource(resources.ModelResource):
    Zoh = fields.Field(column_name='Zoh', attribute='Zoh', widget=ForeignKeyWidget(zoh, field='ZOH'))
    Fab_type = fields.Field(column_name='Fab_type', attribute='Zoh', widget=ForeignKeyWidget(zoh, field='Type'))
    Mill = fields.Field(column_name='Mill', attribute='Mill', widget=ForeignKeyWidget(mill, field='MILL'))
    Ship_to = fields.Field(column_name='Ship_to', attribute='Ship_to', widget=ForeignKeyWidget(factory, field='SHIP_to'))
    
    class Meta:
        model = fabPO
        fields = ('id',"BC_PO_Date","BC_PO_No","Comment","Content","Coo","Cost_Price","Cuttable_Width_cm","Discription","ETA","ETD","Fabric_source","Full_width_cm",
                     "Inv_No","Mill","Mill_Current_Ex_mill","Mill_Orig_Ex_Mill","Mill","Order_Qty","Order_status","PIC","PO_No","PO_Recieved_Date",
                     "Peerless_Current_Ex_mill","Peerless_Orig_Ex_Mill","Peerless_Price","RC_CL_Approval","SO_No","SS_Approval","Season","Ship_to",
                     "Shipped_Qty","Weight_GSM","Yarn_Size","Zoh",'Payment_sts','PI_No','Freight_cost','Fab_type')
        

class PersonDetailAdmin(ImportExportModelAdmin):
    list_display = ('PO_No','Zoh','Mill','Ship_to','Order_Qty','Order_status')
    resource_class = PODetailsAdminResource

class ZOHDetailsAdminResource(resources.ModelResource):
    
    class Meta:
        model = zoh


class ZOHDetailAdmin(ImportExportModelAdmin):
    #list_display = ('id','Zoh')
    resource_class = ZOHDetailsAdminResource



admin.site.register(fabPO,PersonDetailAdmin)
admin.site.register(zoh,ZOHDetailAdmin)
admin.site.register(mill)
admin.site.register(factory)
