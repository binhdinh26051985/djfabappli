from django.db import models
import import_export
from import_export import fields, resources
import datetime
from datetime import datetime


# Create your models here.

class zoh(models.Model):
    fabtype = (
        ('CK Lauren Zhen Xiang','CK Lauren Zhen Xiang'),
        ('GAB open order','GAB open order'),
        ('Open Orders','Open Orders'),
        ('Wool 52 Poly Spandex', 'Wool 52 Poly Spandex'),
        ('Wool Material Open Orders', 'Wool Material Open Orders'),
        ('Wool rich 70 blend polyester', 'Wool rich 70 blend polyester'),
        
        )
    ZOH = models.CharField(null=True,max_length=20, unique=True,db_index=False)
    Type = models.CharField(max_length=200, null=True, choices=fabtype)
    
    def __str__(self):
        return self.ZOH
    
class mill(models.Model):
    MILL = models.CharField(max_length=20, null=True, unique=True)
    Mill_Code = models.CharField(max_length=20, null=True)
    Payment_Term = models.CharField(max_length=20, null=True)
    IPU_mill = models.BooleanField(default=False)
    
    def __str__(self):
        return self.MILL


class factory(models.Model):
    SHIP_to = models.CharField(max_length=100, null=True, unique=True)
    BILL_to = models.CharField(max_length=100, null=True)
    Destination = models.CharField(max_length=100, null=True)
    Customer_code = models.CharField(max_length=100, null=True)
    Deliver_Term = models.CharField(max_length=100, null=True)
    Contact_info = models.CharField(max_length=100, null=True)
    Payment_Term = models.CharField(max_length=100, null=True)
    IPU_fty = models.BooleanField(default=False)
    def __str__(self):
        return self.SHIP_to
    
    
class fabPO(models.Model):
    country = (
        ('CHINA','CHINA'),
        ('INDIA','INDIA'),
        ('TURKEY','TURKEY'),
        ('VIETNAM', 'VIETNAM'),
        ('OTHERS', 'OTHERS')
        )
    
    Ordersts = (
        ('SHIP','SHIP'),
        ('OPEN', 'OPEN')
    )
    
    PO_No = models.CharField(max_length=100)
    Zoh = models.ForeignKey(zoh,null=True, on_delete= models.SET_NULL,blank=True)
    PO_Recieved_Date = models.DateField(null=True, blank=True)
    BC_PO_Date = models.DateField(blank=True, null=True)
    BC_PO_No = models.CharField(max_length=200, null=True,blank=True)
    Discription = models.CharField(max_length=1000, null=True)
    Content = models.CharField(max_length=1000, null=True)
    Order_Qty = models.DecimalField(default=0.00, max_digits=100, decimal_places=2)
    Coo = models.CharField(max_length=200, null=True, choices=country)
    Full_width_cm = models.IntegerField(null=True)
    Cuttable_Width_cm = models.IntegerField(null=True)
    Weight_GSM = models.IntegerField(null=True)
    Yarn_Size = models.CharField(max_length=200, null=True)
    Season = models.CharField(max_length=200, null=True)
    PIC = models.CharField(max_length=200, null=True, blank=True)
    Fabric_source = models.CharField(max_length=200, null=True, blank=True)
    Comment = models.CharField(max_length=2000, null=True)
    Peerless_Price = models.DecimalField(default=0.00, max_digits=100, decimal_places=2)
    Cost_Price = models.DecimalField(default=0.00, max_digits=100, decimal_places=2)
    Freight_cost = models.DecimalField(default=0.00, max_digits=100, decimal_places=2) 
    Mill = models.ForeignKey(mill, null=True, on_delete= models.SET_NULL)
    Mill_Orig_Ex_Mill = models.DateField(null=True, blank=True)
    Mill_Current_Ex_mill = models.DateField(null=True, blank=True)
    Peerless_Orig_Ex_Mill = models.DateField(null=True, blank=True)
    Peerless_Current_Ex_mill = models.DateField(null=True, blank=True)
    SO_No = models.CharField(max_length=200, null=True, blank=True)
    Ship_to = models.ForeignKey(factory, null=True, on_delete= models.SET_NULL,blank=True)
    Shipped_Qty = models.DecimalField(default=0.00, max_digits=100, decimal_places=2)
    SS_Approval = models.CharField(max_length=1000, null=True,blank=True)
    RC_CL_Approval = models.CharField(max_length=1000, null=True,blank=True)
    PI_No = models.CharField(max_length=1000, null=True,blank=True)
    ETD = models.CharField(max_length=1000, null=True,blank=True)
    ETA = models.DateField(null=True, blank=True)
    Inv_No = models.CharField(max_length=1000, null=True, blank=True)
    Payment_sts = models.BooleanField(default=False)
    Order_status = models.CharField(max_length=200, null=True, choices=Ordersts)
    
    
    def __str__(self):
        return str(self.PO_No)


    @property
    
    def Balqty(self):
        x= self.Order_Qty - self.Shipped_Qty
        return x
    
    def Tolerance(self):
        return str( round(((self.Shipped_Qty - self.Order_Qty)/self.Order_Qty * 100),2)) + " %"
    
    def Actual_sale(self):
        return (self.Shipped_Qty * self.Peerless_Price)
    
    
    def Forecast_sale(self):
        
        if (self.Shipped_Qty==0):
            x = round(self.Order_Qty * self.Peerless_Price,2)
            return x
        else:
            return 0
        
        
    def Actual_cost(self):
        return round(self.Shipped_Qty * self.Cost_Price,2)
    
    def Forecast_cost(self):
        
        if (self.Shipped_Qty==0):
            x = round(self.Order_Qty * self.Cost_Price,2)
            return x
        else:
            return 0
        
    def Profit(self):
        
        if (self.Shipped_Qty==0):
            x = self.Order_Qty
        else:
            x = self.Shipped_Qty
            
        return round(x * (self.Peerless_Price - self.Cost_Price),2)
    def CIF(self):
        x = self.Cost_Price + self.Freight_cost
        return x
    
    def Period(self):
        if (self.ETA==None):
            x = str(self.Peerless_Current_Ex_mill.month).zfill(2)
            y = str(self.Peerless_Current_Ex_mill.year)
        else:
            x = self.ETD[5:7]  
            y = self.ETD[:4] 
        return str(y)+ str(x)
        