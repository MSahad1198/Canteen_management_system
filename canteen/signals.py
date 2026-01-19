# canteen/signals.py

from django.db.models.signals import post_save
from django.dispatch import receiver
from .models import Order

@receiver(post_save, sender=Order)
def handle_order_paid(sender, instance, created, **kwargs):
    """
    Reduce stock after order is paid
    and detect low / out of stock products
    """

    # 🚫 Only run when order is paid
    if not instance.is_paid or instance.cancelled:
        return

    low_stock = []
    out_of_stock = []

    for item in instance.items.all():
        product = item.product

        # 🔻 Reduce stock
        product.stock_qty -= item.quantity

        # ❌ Out of stock
        if product.stock_qty <= 0:
            product.stock_qty = 0
            product.in_stock = False
            out_of_stock.append(product.name)

        # ⚠ Low stock (1–5)
        elif product.stock_qty <= 5:
            low_stock.append({
                "name": product.name,
                "remaining": product.stock_qty
            })

        product.save()

    # 🔎 Save alerts temporarily on order (we will read this later)
    instance.low_stock_alert = low_stock
    instance.out_of_stock_alert = out_of_stock
