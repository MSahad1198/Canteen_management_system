from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone

ORDER_TYPE_CHOICES = [
    ('dinein', 'Dine-in'),
    ('takeaway', 'Takeaway'),
]

PAYMENT_METHOD_CHOICES = [
    ('cash', 'Cash'),
    ('card', 'Card'),
    ('wallet', 'Mobile Wallet'),
]

class Category(models.Model):
    name = models.CharField(max_length=100, unique=True)
    description = models.TextField(blank=True)

    def __str__(self):
        return self.name

class Product(models.Model):
    category = models.ForeignKey(Category, on_delete=models.SET_NULL, null=True, blank=True, related_name='products')
    name = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    price = models.DecimalField(max_digits=9, decimal_places=2)
    image = models.ImageField(upload_to='products/', blank=True, null=True)
    in_stock = models.BooleanField(default=True)
    stock_qty = models.IntegerField(default=0)  # inventory count
    tags = models.CharField(max_length=200, blank=True)  # optional tags

    def __str__(self):
        return self.name

class Order(models.Model):
    created_at = models.DateTimeField(auto_now_add=True)
    cashier = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    order_type = models.CharField(max_length=20, choices=ORDER_TYPE_CHOICES, default='takeaway')
    discount_amount = models.DecimalField(max_digits=9, decimal_places=2, default=0)
    is_paid = models.BooleanField(default=False)
    cancelled = models.BooleanField(default=False)
    cancel_reason = models.TextField(blank=True)
    total_amount = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    def __str__(self):
        return f'Order #{self.id} - {self.created_at:%Y-%m-%d %H:%M}'

class OrderItem(models.Model):
    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name='items')
    product = models.ForeignKey(Product, on_delete=models.PROTECT)
    quantity = models.PositiveIntegerField(default=1)
    unit_price = models.DecimalField(max_digits=9, decimal_places=2)

    def line_total(self):
        return self.unit_price * self.quantity

    def __str__(self):
        return f'{self.product.name} x {self.quantity}'

class Payment(models.Model):
    order = models.OneToOneField(Order, on_delete=models.CASCADE, related_name='payment')
    method = models.CharField(max_length=20, choices=PAYMENT_METHOD_CHOICES)
    paid_amount = models.DecimalField(max_digits=10, decimal_places=2)
    change_given = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'Payment for Order #{self.order.id} - {self.method}'

class Combo(models.Model):
    name = models.CharField(max_length=200, unique=True)
    items = models.ManyToManyField(Product, related_name='combos')
    total_price = models.DecimalField(max_digits=9, decimal_places=2)
    discount_amount = models.DecimalField(max_digits=9, decimal_places=2, default=0)
    final_price = models.DecimalField(max_digits=9, decimal_places=2)
    show_in_pos = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    def save(self, *args, **kwargs):
        # ensure final_price is recalculated if needed
        if not self.final_price:
            self.final_price = self.total_price - self.discount_amount
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name} ({self.final_price})"
