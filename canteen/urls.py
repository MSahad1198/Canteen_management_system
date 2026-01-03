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
    path('login/', views.login, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('orders/', views.view_orders, name='view_orders'),
    path('cashier/dashboard/', views.cashier_dashboard, name='cashier_dashboard'),

    # ✅ NEW
    path('combos/', views.manage_combos, name='manage_combos'),
    path('api/create-combo/', views.api_create_combo, name='api_create_combo'),
    path('api/combo-toggle/<int:combo_id>/', views.api_toggle_combo, name='api_toggle_combo'),

    path('receipt/<int:order_id>/', views.reprint_receipt_view, name='reprint_receipt'),

    path('report/', views.report, name='report'),

    path('report/pdf/', views.export_report_pdf, name='report_pdf'),
    path('report/excel/', views.export_report_excel, name='report_excel'),

    path('order/cancel/<int:order_id>/', views.cancel_order, name='cancel_order'),



]

