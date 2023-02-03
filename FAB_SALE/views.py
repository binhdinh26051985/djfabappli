from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse
from django.contrib.auth.models import User, auth
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import *
from .forms import zohform, millform, factoryform, fabPOform
from .filters import OrderFilter
from django.db.models import Q
from django.core.exceptions import FieldDoesNotExist
from django.core.exceptions import ObjectDoesNotExist
import csv
from import_export import resources,fields
from django.db.models import Sum , F , ExpressionWrapper , DecimalField , Case, Value , When, Func
from django.db.models.functions import Round
from django.views.generic import View
from .utils import render_to_pdf
from django.template.loader import get_template
import datetime
from xhtml2pdf import pisa
import xlwt
import pandas as pd
import numpy as np
from io import BytesIO


# Create your views here.
#def homepage(request):
    #return render(request,'fabric.html')

def homepage(request):
    if request.method == "POST":
        username = request.POST['username']
        password = request.POST['password']
        
        user = auth.authenticate(username = username, password = password)
        
        if user is None:
            context = {'Error': 'Invalid username or password.'}
            
            return render(request,"fabric.html", context)
        
            auth.login(request, user)
        else:
            return redirect("home")
    return render(request, "fabric.html")
    

'''

class Exportsale(resources.ModelResource):
    Z_O_H = fields.Field(attribute='Zoh__ZOH')
    Tolerance = fields.Field(attribute='Tolerance')
    Actual_sale = fields.Field(attribute='Actual_sale')
    Forecast_SALE = fields.Field(attribute='Forecast_sale')
    Actual_cost = fields.Field(attribute='Actual_cost')
    Forecast_cost = fields.Field(attribute='Forecast_cost')
    PROFIT = fields.Field(attribute='Profit')
    SHIP_T_O = fields.Field(attribute='Ship_to__SHIP_to')
    IPUFTY = fields.Field(attribute='Ship_to__IPU_fty')
    M_I_L_L = fields.Field(attribute='Mill__MILL')
    IPUMILL = fields.Field(attribute='Mill__IPU_mill')
    
    
    
    class Meta:
        model = fabPO
        fields = ('id','Z_O_H',"Order_status","PO_No","Discription","Content","Peerless_Current_Ex_mill",
                  "Peerless_Orig_Ex_Mill","Order_Qty","Peerless_Price","SHIP_T_O","Shipped_Qty","Inv_No","ETD","ETA",'Tolerance','PI_No','Payment_sts',
                  'Actual_sale','Forecast_SALE','Actual_cost','Forecast_cost','PROFIT','IPUMILL','IPUFTY',
                  "M_I_L_L","Mill_Current_Ex_mill","Mill_Orig_Ex_Mill","Cost_Price",'Freight_cost',"Coo","Full_width_cm","Cuttable_Width_cm","Weight_GSM",
                  "Yarn_Size","PIC","RC_CL_Approval","SS_Approval","Season", "Comment","PO_Recieved_Date",
                  "BC_PO_No","BC_PO_Date","SO_No","Fabric_source")

def export_sale_csv(request):
    fab_Sale = Exportsale()
    dataset = fab_Sale.export()
    response = HttpResponse(dataset.xls, content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Fabric_data.xls"'

    return response
'''

def home(request):
    
    if 'query' in request.GET:
        query=request.GET['query']
        #total_fab_orders = fabPO.objects.filter(PO_No__icontains=query)
        multiple_q = Q(Q(Zoh__ZOH__icontains=query) | Q(PO_No__icontains=query) | Q(Inv_No__icontains=query))
        total_fab_orders = fabPO.objects.filter(multiple_q)
        ord = fabPO.objects.filter(multiple_q).aggregate(
            Order_Qty=Round(Sum('Order_Qty')),
            Ship_Qty=Round(Sum('Shipped_Qty')))
        
        Sale = fabPO.objects.filter(multiple_q).annotate(
            total_actsale = ExpressionWrapper(Round(
                F('Shipped_Qty')*F('Peerless_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Sale_total=Sum('total_actsale'))
    
        Est_Sale = fabPO.objects.filter(multiple_q,Shipped_Qty=0).annotate(
            total_estsale = ExpressionWrapper(
                F('Order_Qty')*F('Peerless_Price'),
                output_field=DecimalField()
                )
            ).aggregate(Est_Sale_TT=Sum('total_estsale'))
    
        Cost = fabPO.objects.filter(multiple_q).annotate(
            total_cost = ExpressionWrapper(Round(
                F('Shipped_Qty')*F('Cost_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Cost_total=Sum('total_cost'))
    
        Est_Cost = fabPO.objects.filter(multiple_q,Shipped_Qty=0).annotate(
            total_Cost = ExpressionWrapper(Round(
                F('Order_Qty')*F('Cost_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Est_Cost_TT=(Sum('total_Cost')),)
        
        
        GM = fabPO.objects.filter(multiple_q).aggregate(
            amount=Round(Sum(Case(
                When(Shipped_Qty=0, then= F('Order_Qty')*(F('Peerless_Price') - F('Cost_Price'))), 
                default=F('Shipped_Qty')* (F('Peerless_Price') - F('Cost_Price' ))),
                output_field=DecimalField()
                 ))
            )
        
        
    else:
        total_fab_orders = fabPO.objects.filter(PO_No__icontains="UNKNOWN")
        ord = fabPO.objects.all().aggregate(
            Order_Qty=Round(Sum('Order_Qty')),
            Ship_Qty=Round(Sum('Shipped_Qty')))
        
        Sale = fabPO.objects.all().annotate(
            
            total_actsale = ExpressionWrapper(Round(
                F('Shipped_Qty')*F('Peerless_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Sale_total=Sum('total_actsale'))
    
        Est_Sale = fabPO.objects.filter(Shipped_Qty=0).annotate(
            total_estsale = ExpressionWrapper(
                F('Order_Qty')*F('Peerless_Price'),
                output_field=DecimalField()
                )
            ).aggregate(Est_Sale_TT=Sum('total_estsale'))
    
    
        Cost = fabPO.objects.all().annotate(
            total_cost = ExpressionWrapper(Round(
                F('Shipped_Qty')*F('Cost_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Cost_total=Sum('total_cost'))
    
    
        Est_Cost = fabPO.objects.filter(Shipped_Qty=0).annotate(
            total_Cost = ExpressionWrapper(Round(
                F('Order_Qty')*F('Cost_Price')),
                output_field=DecimalField()
                )
            ).aggregate(Est_Cost_TT=(Sum('total_Cost')))
    
    
        GM = fabPO.objects.all().aggregate(
            amount=Round(Sum(Case(
                When(Shipped_Qty=0, then=(F('Order_Qty')*(F('Peerless_Price') - F('Cost_Price')))), 
                default=(F('Shipped_Qty')* (F('Peerless_Price') - F('Cost_Price' )))),
                output_field=DecimalField()
                 ))
            )

    context = {'total_fab_orders': total_fab_orders,
               'ord':ord,
               'Sale':Sale,
               'Est_Sale':Est_Sale,
               'Cost':Cost,
               'Est_Cost':Est_Cost,
               'GM': GM
               }
    
    return render(request, "index.html", context)

def createorder(request):
	form = fabPOform()
	if request.method == 'POST':
		#print('Printing POST:', request.POST)
		form = fabPOform(request.POST)
		if form.is_valid():
			form.save()
			return redirect("home")

	context = {'form':form}
	return render(request, 'neworderform.html', context)


def updateorder(request, pk):

	order = fabPO.objects.get(id=pk)
	form = fabPOform(instance=order)

	if request.method == 'POST':
		form = fabPOform(request.POST, instance=order)
		if form.is_valid():
			form.save()
			return redirect('home')

	context = {'form':form}
	return render(request, 'neworderform.html', context)


def Createzoh(request):
	form = zohform()
	if request.method == 'POST':
		#print('Printing POST:', request.POST)
		form = zohform(request.POST)
		if form.is_valid():
			form.save()
			return redirect('home')

	context = {'form':form}
	return render(request, 'Zohform.html', context)


def Createmill(request):
	form = millform()
	if request.method == 'POST':
		#print('Printing POST:', request.POST)
		form = millform(request.POST)
		if form.is_valid():
			form.save()
			return redirect('home')

	context = {'form':form}
	return render(request, 'millform.html', context)

def Createfactory(request):
	form = factoryform()
	if request.method == 'POST':
		#print('Printing POST:', request.POST)
		form = factoryform(request.POST)
		if form.is_valid():
			form.save()
			return redirect('home')

	context = {'form':form}
	return render(request, 'factoryform.html', context)


def export_file(request):
    #multiple_q = Q(Q(Ship_to__IPU_fty=False) & (~Q(Ship_to__SHIP_to="TO BE CONFIRM")))
    
    data = fabPO.objects.all()
    #data = fabPO.objects.filter(multiple_q)
    
    response = HttpResponse('text/csv')
    response['Content-Disposition'] = 'attachment; filename=Fabric.csv'
    writer = csv.writer(response)
    writer.writerow(['id',"Order_status","Zoh__ZOH","PO_No","Discription","Content","Peerless_Current_Ex_mill","Peerless_Orig_Ex_Mill","Order_Qty","Peerless_Price",
                             "Ship_to__SHIP_to","Shipped_Qty","Inv_No","ETD","ETA","Mill__MILL","Mill_Current_Ex_mill","Mill_Orig_Ex_Mill","Cost_Price",
                             'Freight_cost','PI_No','Payment_sts','Ship_to__Payment_Term','Ship_to__IPU_fty','Mill__Payment_Term','Mill__IPU_mill',
                             "Coo","Full_width_cm","Cuttable_Width_cm","Weight_GSM","Yarn_Size","PIC","RC_CL_Approval","SS_Approval","Season",
                            "Comment","PO_Recieved_Date","BC_PO_No","BC_PO_Date","SO_No","Fabric_source",'Zoh__Type'  
                    ])
    data1 = data.values_list('id',"Order_status","Zoh__ZOH","PO_No","Discription","Content","Peerless_Current_Ex_mill","Peerless_Orig_Ex_Mill","Order_Qty","Peerless_Price",
                             "Ship_to__SHIP_to","Shipped_Qty","Inv_No","ETD","ETA","Mill__MILL","Mill_Current_Ex_mill","Mill_Orig_Ex_Mill","Cost_Price",
                             'Freight_cost','PI_No','Payment_sts','Ship_to__Payment_Term','Ship_to__IPU_fty','Mill__Payment_Term','Mill__IPU_mill',
                             "Coo","Full_width_cm","Cuttable_Width_cm","Weight_GSM","Yarn_Size","PIC","RC_CL_Approval","SS_Approval","Season",
                            "Comment","PO_Recieved_Date","BC_PO_No","BC_PO_Date","SO_No","Fabric_source",'Zoh__Type')
    
    for Dat in data1:
        writer.writerow(Dat)
    return response


def PI(request):
        
    if 'query' in request.GET:
        query=request.GET['query']
        total_fab_orders = fabPO.objects.filter(PI_No__icontains=query) 
        
        ordpi = total_fab_orders.aggregate(
            Ship_Qty=Sum('Shipped_Qty'))
        
        Salepi = total_fab_orders.annotate(
            total_actsale = ExpressionWrapper(Round(
                F('Shipped_Qty')*F('Peerless_Price')*100)/100,
                output_field=DecimalField(max_digits=12, decimal_places=2)
                )
            ).aggregate(Sale_total=Sum('total_actsale'))
        
        ft1 = total_fab_orders.values_list('Ship_to__SHIP_to', flat = True)
        fty = factory.objects.filter(SHIP_to__icontains=ft1)
        
        if total_fab_orders.values_list('PI_No', flat = True).exists():
            PI1 = total_fab_orders.values_list('PI_No', flat = True)[0]
        else:
            PI1 = {'0': 0}
            return PI1
    else:
        total_fab_orders = {'0':0}
        ordpi = {'0':0}
        Salepi = {'0':0}
        ft1 = {'0':0}
        fty= {'0':0}
        PI1= {'0':0}
   
    issuedate = datetime.date.today()
    
    context = {'total_fab_orders': total_fab_orders,
               'ordpi':ordpi,
               'Salepi':Salepi,
               'fty': fty,
               "PI1": PI1,
               "issuedate": issuedate,
               }
    return render(request, "PI.html", context)

@login_required
def searchPI(request):
        
    if 'query' in request.GET:
        
        query=request.GET['query']
        total_fab_orders = fabPO.objects.filter(PI_No=query)
        PIsum = total_fab_orders.values('Ship_to__SHIP_to','Ship_to__BILL_to','Ship_to__Destination','Ship_to__Customer_code','Ship_to__Deliver_Term','Ship_to__Contact_info',
                                    'Ship_to__Payment_Term','PI_No','PO_No','Zoh__ZOH','Discription','Content','Shipped_Qty','Peerless_Price','Cuttable_Width_cm',
                                    'Peerless_Current_Ex_mill')
        PINo = query
        
        
    else:
        total_fab_orders = fabPO.objects.filter(PI_No__icontains="UNKNOWN")
        total_fab_order1 = fabPO.objects.all()
        PINo = "NOT FOUND"
        
        PIsum = total_fab_order1.values('Ship_to__SHIP_to','Ship_to__BILL_to','Ship_to__Destination','Ship_to__Customer_code','Ship_to__Deliver_Term','Ship_to__Contact_info',
                                    'Ship_to__Payment_Term','PI_No','PO_No','Zoh__ZOH','Discription','Content','Shipped_Qty','Peerless_Price','Cuttable_Width_cm',
                                    'Peerless_Current_Ex_mill')
    data = pd.DataFrame(PIsum)
    data['Amount$'] = (data['Shipped_Qty']* data['Peerless_Price'])
    
    html_table = data.to_html(index=False)
    
    summary = data.pivot_table(index=['PI_No'],
                             values=['Shipped_Qty','Amount$'],
                             aggfunc=np.sum, margins=True)
    html_summary = summary.to_html(index=True)
    
    
    
    issuedate = datetime.date.today()
    
    context = {'issuedate':issuedate,
               'html_table': html_table,
               'html_summary':html_summary,
               'PINo': PINo,
               }
    
    return render(request, "Search PI.html", context)


def detail_PI(request,pk):
    if fabPO.objects.filter(PI_No=pk).exists():
        total_fab_orders = fabPO.objects.filter(PI_No=pk)
    else:
        total_fab_orders = fabPO.objects.all()
    
    PIsum = total_fab_orders.values('Ship_to__SHIP_to','Ship_to__BILL_to','Ship_to__Destination','Ship_to__Customer_code','Ship_to__Deliver_Term','Ship_to__Contact_info',
                                    'Ship_to__Payment_Term','PI_No','PO_No','BC_PO_No','Zoh__ZOH','Discription','Content','Shipped_Qty','Order_Qty','Peerless_Price','Cuttable_Width_cm',
                                    'Peerless_Current_Ex_mill')
    data = pd.DataFrame(PIsum)
    
    data['Amount$'] = data['Shipped_Qty'] * data['Peerless_Price']
    
    customer = data[['Ship_to__SHIP_to','Ship_to__BILL_to','Ship_to__Destination','Ship_to__Customer_code','Ship_to__Deliver_Term','Ship_to__Contact_info',
                                    'Ship_to__Payment_Term']]
    customer.columns = ['Ship to:','Bill to:','Destination:','Customer code:','Delivery Term:','Contact info:','Payment Term:']
    customer = customer[['Ship to:','Bill to:','Destination:','Customer code:','Delivery Term:','Contact info:','Payment Term:']].iloc[0]
    
    
    customer = customer.reset_index()
    customer.columns = ['Description','Destination']
    
    cust_html = customer.to_html(index=False)
    
    data1 = data[['PO_No','BC_PO_No','Zoh__ZOH','Discription','Content','Shipped_Qty','Peerless_Price','Cuttable_Width_cm','Peerless_Current_Ex_mill','Amount$']]
    data1.columns = ['PO_No','BC_No','ZROH','Ref','Content','Quantity','Price $/m','Cut_width_cm','Ex_mill','Amount$']
    html_table = data1.to_html(index=False)
    
    summary = data.pivot_table(index=['PI_No'],
                             values=['Shipped_Qty','Amount$'],
                             aggfunc=np.sum, margins=False)
    
    html_summary = summary.to_html(index=False)
    PInumber = pk
    
    issuedate = datetime.date.today()
    context = {'issuedate':issuedate,
               'html_table': html_table,
               'html_summary':html_summary,
               'PInumber': PInumber,
               'cust_html':cust_html,
               }
    
    
    #return render(request,'PI_Export.html',context) 
    
    #template = get_template('PI_Export.html')
    #html = template.render(context)
    #pdf = render_to_pdf('PI_Export.html', context)
    #return (pdf)
    
    #return render_to_pdf('PI_Export.html',context)
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'filename="PI.pdf"'
    template = get_template('PI_Export.html')
    html = template.render(context)
    pisa_status = pisa.CreatePDF(html, dest=response)
    if pisa_status.err:
        return HttpResponse('errors <pre>' + html + '</pre>')
    return response
    #return render(request,'PI_Export.html',context)
    #return HttpResponse(pdf,content_type='application/pdf')

    


def salereport(request):
    multiple_q = Q(Q(Ship_to__IPU_fty=False) & (~Q(Ship_to__SHIP_to="TO BE CONFIRM")))
    saleorder = fabPO.objects.filter(multiple_q)
    salesum = saleorder.values('BC_PO_Date', 'BC_PO_No','Content','Cost_Price','Discription', 'ETA', 'ETD','Inv_No','Order_Qty','PO_No','Payment_sts',
                               'Peerless_Current_Ex_mill',
                               'Peerless_Price','Ship_to__SHIP_to','Shipped_Qty','Zoh__ZOH')
    data = pd.DataFrame(salesum)
    data['1.Act_Sale'] = (data['Shipped_Qty']* data['Peerless_Price'])
    data['2.Est_Sale'] = np.where(data['Shipped_Qty'] ==0,data['Order_Qty']*data['Peerless_Price'],0)
    data['3.Total_Sale'] = data['1.Act_Sale']+ data['2.Est_Sale']
    data['Month'] = np.where(data['ETA'].isnull(),data['Peerless_Current_Ex_mill'].astype(str).str[:7],data['ETD'].str[:7])
    data['Year'] = data['Month'].str[:4]
    html_table = data.to_html(index=False)
    
    summary = data.pivot_table(index=['Year',"Month"],
                             values=['Order_Qty','Shipped_Qty','1.Act_Sale','2.Est_Sale','3.Total_Sale'],
                             aggfunc=np.sum, margins=True)
    html_summary = summary.to_html(index=True)
    
    
    
    issuedate = datetime.date.today()
    context = {'issuedate':issuedate,
               'html_table': html_table,
               'html_summary':html_summary,
               }
    return render(request,'salereport.html',context)  

def salereportexport(request):
    multiple_q = Q(Q(Ship_to__IPU_fty=False) & (~Q(Ship_to__SHIP_to="TO BE CONFIRM")))
    saleorder = fabPO.objects.filter(multiple_q)
    salesum = saleorder.values('BC_PO_Date', 'BC_PO_No','Content','Cost_Price','Discription', 'ETA', 'ETD','Inv_No','Order_Qty','PO_No','Payment_sts',
                               'Peerless_Current_Ex_mill',
                               'Peerless_Price','Ship_to__SHIP_to','Shipped_Qty','Zoh__ZOH')
    data = pd.DataFrame(salesum)
    data['1.Act_Sale'] = (data['Shipped_Qty']* data['Peerless_Price'])
    data['2.Est_Sale'] = np.where(data['Shipped_Qty'] ==0,data['Order_Qty']*data['Peerless_Price'],0)
    data['3.Total_Sale'] = data['1.Act_Sale']+ data['2.Est_Sale']
    data['Month'] = np.where(data['ETA'].isnull(),data['Peerless_Current_Ex_mill'].astype(str).str[:7],data['ETD'].str[:7])
    data['Year'] = data['Month'].str[:4]
    data1 = data[['PO_No','BC_PO_No','Zoh__ZOH','Content','Order_Qty','Peerless_Price','Peerless_Current_Ex_mill','Ship_to__SHIP_to','ETD','ETA','Shipped_Qty','Inv_No',
                 '1.Act_Sale','2.Est_Sale','3.Total_Sale','Month','Year']]
    
    summary = data.pivot_table(index=['Year',"Month"],
                             values=['Order_Qty','Shipped_Qty','1.Act_Sale','2.Est_Sale','3.Total_Sale'],
                             aggfunc=np.sum, margins=True)
    
    
    
    with BytesIO() as b:
        
        writer = pd.ExcelWriter(b, engine='xlsxwriter')
        data1.to_excel(writer, sheet_name='Data',index=False)
        summary.to_excel(writer, sheet_name='Sum')
        writer.save()
        # Set up the Http response.
        filename = 'Fabric Sale.xlsx'
        response = HttpResponse(
            b.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename
        return response
        
    
def Purchase(request):
    #multiple_q = Q(Q(Ship_to__IPU_fty=True) & (~Q(Ship_to__SHIP_to="TO BE CONFIRM")))
    saleorder = fabPO.objects.filter(Ship_to__SHIP_to="DAESE")
    salesum = saleorder.values('Mill__MILL','PO_No','BC_PO_No','Zoh__ZOH','Content','Peerless_Current_Ex_mill','Ship_to__SHIP_to','Order_Qty','Shipped_Qty','Cost_Price',
                               'Freight_cost','Inv_No','ETD','ETA')
    data = pd.DataFrame(salesum)
    data['Qty'] = np.where(data['Shipped_Qty']==0,data['Order_Qty'],data['Shipped_Qty'])
    data['Price'] = data['Cost_Price']+data['Freight_cost']
    data['1.Act_Amount'] = data['Shipped_Qty']* data['Price']
    data['2.Est_Amount'] = np.where(data['Shipped_Qty'] ==0,data['Order_Qty']*data['Price'],0)
    data['3.Total_Amount'] = data['1.Act_Amount']+ data['2.Est_Amount']
    data['Month'] = np.where(data['ETA'].isnull(),data['Peerless_Current_Ex_mill'].astype(str).str[:7],data['ETD'].str[:7])
    data['Year'] = data['Month'].str[:4]
    data1 = data[['Mill__MILL','PO_No','BC_PO_No','Zoh__ZOH','Content','Peerless_Current_Ex_mill','Ship_to__SHIP_to','Qty','Price','1.Act_Amount','2.Est_Amount','3.Total_Amount','Inv_No','ETD','Month','Year']]
    html_table = data1.to_html(index=False)
    
    summary = data.pivot_table(index=['Year',"Month"],
                             values=['Qty','1.Act_Amount','2.Est_Amount','3.Total_Amount'],
                             aggfunc=np.sum, margins=True)
    html_summary = summary.to_html(index=True)
    
    
    
    issuedate = datetime.date.today()
    context = {'issuedate':issuedate,
               'html_table': html_table,
               'html_summary':html_summary,
               }
    return render(request,'Purchasereport.html',context)

def purchasexport(request):
    saleorder = fabPO.objects.filter(Ship_to__SHIP_to="DAESE")
    salesum = saleorder.values('Mill__MILL','PO_No','BC_PO_No','Zoh__ZOH','Content','Peerless_Current_Ex_mill','Ship_to__SHIP_to','Order_Qty','Shipped_Qty','Cost_Price',
                               'Freight_cost','Inv_No','ETD','ETA')
    data = pd.DataFrame(salesum)
    data['Mill_name'] = "IPU_MILL-"+data['Mill__MILL']
    data['Qty'] = np.where(data['Shipped_Qty']==0,data['Order_Qty'],data['Shipped_Qty'])
    data['Price'] = data['Cost_Price']+data['Freight_cost']
    data['1.Act_Amount'] = data['Shipped_Qty']* data['Price']
    data['2.Est_Amount'] = np.where(data['Shipped_Qty'] ==0,data['Order_Qty']*data['Price'],0)
    data['3.Total_Amount'] = data['1.Act_Amount']+ data['2.Est_Amount']
    data['Month'] = np.where(data['ETA'].isnull(),data['Peerless_Current_Ex_mill'].astype(str).str[:7],data['ETD'].str[:7])
    data['Year'] = data['Month'].str[:4]
    datacol = ['Qty','Price','1.Act_Amount','2.Est_Amount','3.Total_Amount']
    data[datacol]=data[datacol].apply(pd.to_numeric).fillna(0)
    
    data1 = data[['Mill_name','PO_No','BC_PO_No','Zoh__ZOH','Content','Peerless_Current_Ex_mill','Ship_to__SHIP_to','Qty','Price','1.Act_Amount','2.Est_Amount','3.Total_Amount','Inv_No','ETD','Month','Year']]
    
    
    sum1 = data[['Year',"Month",'Qty','1.Act_Amount','2.Est_Amount','3.Total_Amount']]
    
    
    summary = sum1.groupby(['Year',"Month"]).sum()
    df1 = summary.groupby(["Year"]).sum()
    df1.index = [df1.index.get_level_values(0),
                #df1.index.get_level_values(1),
                ['Total'] * len(df1)]
    
    df = pd.concat([summary, df1]).sort_index(level=[0,1])
    

    
    #summary = data.pivot_table(index=['Year',"Month"],
                             #values=['Qty','1.Act_Amount','2.Est_Amount','3.Total_Amount'],
                             #aggfunc=np.sum, margins=True)
    
    with BytesIO() as b:
        # Use the StringIO object as the filehandle.
        writer = pd.ExcelWriter(b, engine='xlsxwriter')
        #data1.to_excel(writer, sheet_name='Data',index=False)
        #summary.to_excel(writer, sheet_name='Sum')
        
        data1.to_excel(writer, sheet_name='Data',index=False)
        df.to_excel(writer, sheet_name='Sum')
        workbook = writer.book
        worksheet1 = writer.sheets['Data']
        worksheet2 = writer.sheets['Sum']
        
        total_fmt = workbook.add_format({'align': 'center', 'num_format': '#,###.00',
                                 'bold': True, 'bottom':1, 'border':1})
        
        worksheet1.set_column("A:P", 12, total_fmt)
        worksheet2.set_column("A:F", 12, total_fmt)
        
        
        writer.save()
        # Set up the Http response.
        filename = 'Fabric Purchase.xlsx'
        response = HttpResponse(
            b.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename
        return response




















    



    
    
    
    
    
    
    

