from django.urls import path
from . import views
from django.contrib.auth.decorators import login_required

app_name = 'canteen'

urlpatterns = [
    path('', views.index, name='index'),
    path('pos/', views.pos_page, name='pos'),
    path('api/products/', views.api_products, name='api_products'),
    path('api/create-order/', views.api_create_order, name='api_create_order'),
    path('api/reprint/<int:order_id>/', views.api_reprint_receipt, name='api_reprint'),
    path('reports/daily/', views.daily_sales_report, name='daily_sales'),
    #path('login/', views.login, name='login'),
   path('logout/', views.logout_view, name='logout'),
    path('orders/', views.view_orders, name='view_orders'),
    path('cashier/dashboard/', views.cashier_dashboard, name='cashier_dashboard'),

    # ✅ NEW
    path('combos/', views.manage_combos, name='manage_combos'),
    path('api/create-combo/', views.api_create_combo, name='api_create_combo'),
    path('api/combo-toggle/<int:combo_id>/', views.api_toggle_combo, name='api_toggle_combo'),

    path('receipt/<int:order_id>/', views.reprint_receipt_view, name='reprint_receipt'),

    path('report/', views.report, name='report'),

    #detasiled reports urls
    path('report/today-sales/', views.today_sales_detail, name='today_sales_detail'),
    path('report/weekly-sales/', views.weekly_sales_detail, name='weekly_sales_detail'),
    path('report/monthly-sales/', views.monthly_sales_detail, name='monthly_sales_detail'),
    path('report/yearly-sales/', views.yearly_sales_detail, name='yearly_sales_detail'),

    path('report/pdf/', views.export_report_pdf, name='report_pdf'),
    path('report/excel/', views.export_report_excel, name='report_excel'),

    path('order/cancel/<int:order_id>/', views.cancel_order, name='cancel_order'),

     # ✅ TODAY'S SALES EXPORTS
    path('report/today-sales/pdf/', views.export_today_sales_pdf, name='export_today_pdf'),
    path('report/today-sales/excel/', views.export_today_sales_excel, name='export_today_excel'),
    
    # ✅ WEEKLY SALES EXPORTS
    path('report/weekly-sales/pdf/', views.export_weekly_sales_pdf, name='export_weekly_pdf'),
    path('report/weekly-sales/excel/', views.export_weekly_sales_excel, name='export_weekly_excel'),
    
    # ✅ MONTHLY SALES EXPORTS
    path('report/monthly-sales/pdf/', views.export_monthly_sales_pdf, name='export_monthly_pdf'),
    path('report/monthly-sales/excel/', views.export_monthly_sales_excel, name='export_monthly_excel'),
    
    # ✅ YEARLY SALES EXPORTS
    path('report/yearly-sales/pdf/', views.export_yearly_sales_pdf, name='export_yearly_pdf'),
    path('report/yearly-sales/excel/', views.export_yearly_sales_excel, name='export_yearly_excel'),

]

