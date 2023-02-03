from django.urls import path
from . import views


urlpatterns = [
    path('', views.homepage, name = 'homepage'),
    path('home/', views.home, name = 'home'),
    path('createorder/', views.createorder, name="createorder"),
    path('updateorder/<str:pk>/', views.updateorder, name="updateorder"),
    path('Createzoh/', views.Createzoh, name="Createzoh"),
    path('Createmill/', views.Createmill, name="Createmill"),
    path('Createfactory/', views.Createfactory, name="Createfactory"),
    path('export_file/', views.export_file, name="export_file"),
    path('PI/', views.PI, name="PI"),
    path('searchPI/', views.searchPI, name="searchPI"),
    #path('<str:pk>', views.detail_page, name="detail"),
    path('salereport/', views.salereport, name = 'salereport'),
    path('salereportexport/', views.salereportexport, name="salereportexport"),
    path('Purchase/', views.Purchase, name = 'Purchase'),
    path('purchasexport/', views.purchasexport, name = 'purchasexport'),
    path('<str:pk>/', views.detail_PI, name = 'detail_PI'),
    #path('<str:pk>', views.detail_PI, name = 'detail_PI'),
    ]
    
    
    #path('generateinvoice/<int:pk>/', views.GenerateInvoice.as_view(), name = 'generateinvoice'),
    #path('GeneratePdf/', views.GeneratePdf, name = 'GeneratePdf'),
