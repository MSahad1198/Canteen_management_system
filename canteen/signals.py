from django.db.models.signals import post_save
from django.dispatch import receiver
from .models import Order, OrderItem, Product

@receiver(post_save, sender=Order)
def handle_order_paid(sender, instance: Order, created, **kwargs):
    """
    When an order becomes paid (is_paid=True), decrement stock.
    """
    # Only handle when order marked paid and not cancelled
    if instance.is_paid and not instance.cancelled:
        for item in instance.items.all():
            product = item.product
            # reduce stock if tracked
            if product.stock_qty is not None:
                product.stock_qty = max(0, product.stock_qty - item.quantity)
                if product.stock_qty == 0:
                    product.in_stock = False
                product.save()
