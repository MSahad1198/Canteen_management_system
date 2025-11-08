from django.contrib import admin
from .models import Category, Product, Order, OrderItem, Payment
from .models import Combo

@admin.register(Combo)
class ComboAdmin(admin.ModelAdmin):
    list_display = ('name', 'final_price', 'show_in_pos', 'created_at')
    list_editable = ('show_in_pos',)
    filter_horizontal = ('items',)


class OrderItemInline(admin.TabularInline):
    model = OrderItem
    extra = 0

@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ('id', 'created_at', 'cashier', 'total_amount', 'is_paid', 'cancelled')
    inlines = [OrderItemInline]
    list_filter = ('is_paid', 'cancelled', 'created_at')

@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ('name', 'category', 'price', 'stock_qty', 'in_stock')
    list_filter = ('category', 'in_stock')



admin.site.register(Category)
admin.site.register(Payment)
