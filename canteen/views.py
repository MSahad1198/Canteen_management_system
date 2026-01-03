import json
from decimal import Decimal
from django.shortcuts import render, get_object_or_404
from django.http import JsonResponse, HttpResponseBadRequest
from django.views.decorators.http import require_POST
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from .models import Product, Order, OrderItem, Payment
from django.contrib.auth import authenticate, login, logout, login as auth_login, logout as auth_logout       # ✅ for authentication
from django.contrib import messages                                # ✅ for flash messages
from django.contrib.auth.forms import AuthenticationForm          # ✅ built-in login form
from django.shortcuts import redirect
from .models import Order
from django.contrib.auth.models import Group
from django.contrib.auth.decorators import user_passes_test
from django.db.models import Sum, Count
from django.contrib.admin.views.decorators import staff_member_required
from .models import Combo, Product
from datetime import timedelta
#exort to pdf import 
from django.http import HttpResponse
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from io import BytesIO
from django.db.models.functions import ExtractYear



@login_required
def report(request):
    return render(request, 'canteen/report.html')

def is_admin(user):
    """Return True if user belongs to the 'Admin' group."""
    return user.groups.filter(name='Admin').exists()

def is_cashier(user):
    """Return True if user belongs to the 'Cashier' group."""
    return user.groups.filter(name='Cashier').exists()

@login_required
def index(request):
    is_admin = request.user.is_authenticated and request.user.groups.filter(name='Admin').exists()
    is_cashier = request.user.is_authenticated and request.user.groups.filter(name='Cashier').exists()
    return render(request, 'canteen/index.html', {
        'is_admin': is_admin,
        'is_cashier': is_cashier,
    })



@login_required
def pos_page(request):
    # POS frontend — will request products via AJAX
    return render(request, 'canteen/pos.html')

def api_products(request):
    qs = Product.objects.filter(in_stock=True).order_by('name')
    combos = Combo.objects.filter(show_in_pos=True)
    data = []

    for p in qs:
        data.append({
            'id': p.id,
            'name': p.name,
            'price': str(p.price),
            'stock_qty': p.stock_qty,
            'type': 'product',
        })
    for c in combos:
        data.append({
            'id': f"combo-{c.id}",
            'name': f"🍱 {c.name}",
            'price': str(c.final_price),
            'stock_qty': 'Combo',
            'type': 'combo',
        })
    return JsonResponse({'products': data})


@require_POST
@login_required
def api_create_order(request):
    try:
        data = json.loads(request.body.decode('utf-8'))
        items = data.get('items', [])

        if not items:
            return HttpResponseBadRequest('No items provided')

        # ✅ Step 1: Validate stock first (before creating the order)
        insufficient = []  # list to collect items with low stock

        for it in items:
            pid = str(it['product_id'])
            qty = int(it.get('quantity', 1))

            if pid.startswith("combo-"):
                combo_id = pid.split("-")[1]
                combo = get_object_or_404(Combo, pk=combo_id)
                for prod in combo.items.all():
                    if prod.stock_qty < qty:
                        insufficient.append(f"{prod.name} (in {combo.name})")
            else:
                product = get_object_or_404(Product, pk=int(pid))
                if product.stock_qty < qty:
                    insufficient.append(product.name)

        # ❌ Stop if any product lacks stock
        if insufficient:
            return JsonResponse({
                "status": "error",
                "message": "Insufficient stock for: " + ", ".join(insufficient)
            }, status=400)

        # ✅ Step 2: Proceed to order creation
        total = Decimal('0')
        order = Order.objects.create(
            cashier=request.user,
            order_type=data.get('order_type', 'takeaway'),
            discount_amount=Decimal(str(data.get('discount_amount', '0'))),
            is_paid=False
        )

        # ✅ Step 3: Add items to order
        for it in items:
            pid = str(it['product_id'])
            qty = int(it.get('quantity', 1))

            if pid.startswith("combo-"):
                combo_id = pid.split("-")[1]
                combo = get_object_or_404(Combo, pk=combo_id)
                for prod in combo.items.all():
                    OrderItem.objects.create(
                        order=order,
                        product=prod,
                        quantity=qty,
                        unit_price=prod.price
                    )
                    total += prod.price * qty
            else:
                product = get_object_or_404(Product, pk=int(pid))
                OrderItem.objects.create(
                    order=order,
                    product=product,
                    quantity=qty,
                    unit_price=product.price
                )
                total += product.price * qty

        total_after_discount = total - order.discount_amount
        order.total_amount = total_after_discount
        order.save()

        # ✅ Step 4: Process payment if provided
        payment_data = data.get('payment')
        if payment_data:
            paid_amount = Decimal(str(payment_data.get('paid_amount', '0')))
            change = Decimal('0')
            if payment_data.get('method') == 'cash':
                change = (paid_amount - total_after_discount) if paid_amount > total_after_discount else Decimal('0')

            Payment.objects.create(
                order=order,
                method=payment_data.get('method', 'cash'),
                paid_amount=paid_amount,
                change_given=change
            )

            order.is_paid = True
            order.save()  # stock updates trigger from signals

        return JsonResponse({
            'status': 'success',
            'order_id': order.id,
            'total': str(order.total_amount),
            'is_paid': order.is_paid
        })

    except Exception as e:
        print("⚠️ Error creating order:", e)
        return JsonResponse({'status': 'error', 'message': str(e)}, status=500)


@login_required
def api_reprint_receipt(request, order_id):
    order = get_object_or_404(Order, pk=order_id)
    items = [{
        'product': it.product.name,
        'qty': it.quantity,
        'unit_price': str(it.unit_price),
        'line_total': str(it.line_total())
    } for it in order.items.all()]
    data = {
        'order_id': order.id,
        'created_at': order.created_at.isoformat(),
        'cashier': order.cashier.username if order.cashier else None,
        'items': items,
        'total': str(order.total_amount)
    }
    return JsonResponse(data)

@login_required
@user_passes_test(is_admin)
def daily_sales_report(request):
    # basic daily report for today
    today = timezone.localdate()
    orders = Order.objects.filter(created_at__date=today, is_paid=True, cancelled=False)
    total = sum(o.total_amount for o in orders)
    best_sellers = {}
    for o in orders:
        for it in o.items.all():
            best_sellers.setdefault(it.product.name, 0)
            best_sellers[it.product.name] += it.quantity
    # convert to sorted list
    best = sorted(best_sellers.items(), key=lambda x: -x[1])[:10]
    return JsonResponse({
        'date': str(today),
        'orders_count': orders.count(),
        'total_sales': str(total),
        'best_sellers': best
    })


def login(request):
    """
    Login page with role selection (Admin or Cashier).
    Admin → Django Admin panel
    Cashier → POS page
    """
    if request.user.is_authenticated:
        return redirect('canteen:index')

    if request.method == 'POST':
        role = request.POST.get('role')
        form = AuthenticationForm(request, data=request.POST)

        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)

            if user is not None:
                auth_login(request, user)

                # ✅ Add user to selected role group (optional)
                if role == 'admin':
                    group, _ = Group.objects.get_or_create(name='Admin')
                    user.groups.add(group)
                    return redirect('canteen:index')  # Django admin interface

                elif role == 'cashier':
                    group, _ = Group.objects.get_or_create(name='Cashier')
                    user.groups.add(group)
                    return redirect('canteen:cashier_dashboard')


                else:
                    return redirect('canteen:index')
            else:
                messages.error(request, "Invalid username or password.")
        else:
            messages.error(request, "Invalid credentials.")
    else:
        form = AuthenticationForm()

    return render(request, 'canteen/login.html', {'form': form})


def logout_view(request):
    from django.contrib.auth import logout
    logout(request)
    return redirect('login')


@login_required
def view_orders(request):
    orders = Order.objects.all().order_by('-created_at')
    return render(request, 'canteen/view_orders.html', {'orders': orders})

def is_cashier(user):
    """Return True if user belongs to 'Cashier' group."""
    return user.groups.filter(name='Cashier').exists()

def is_admin(user):
    """Return True if user belongs to 'Admin' group."""
    return user.groups.filter(name='Admin').exists()

@login_required
def cashier_dashboard(request):
    """Dashboard for cashiers showing their sales summary and recent orders."""
    from .models import Order

    today = timezone.localdate()

    # Filter only this cashier's orders for today
    orders_today = Order.objects.filter(
        cashier=request.user,
        created_at__date=today
    )

    total_sales = orders_today.filter(is_paid=True).aggregate(total=Sum('total_amount'))['total'] or 0
    total_orders = orders_today.count()
    pending_orders = orders_today.filter(is_paid=False).count()

    # Show the last 10 orders
    recent_orders = orders_today.order_by('-created_at')[:10]

    context = {
        'total_sales': f"{total_sales:.2f}",
        'total_orders': total_orders,
        'pending_orders': pending_orders,
        'recent_orders': recent_orders,
    }
    return render(request, 'canteen/cashier_dashboard.html', context)

@user_passes_test(is_admin)
@login_required
def manage_combos(request):
    products = Product.objects.filter(in_stock=True)
    combos = Combo.objects.all().order_by('-created_at')
    return render(request, 'canteen/manage_combos.html', {
        'products': products,
        'combos': combos,
    })


@require_POST
@user_passes_test(is_admin)
@login_required
def api_create_combo(request):
    data = json.loads(request.body)
    combo = Combo.objects.create(
        name=data['name'],
        total_price=data['total_price'],
        discount_amount=data['discount_amount'],
        final_price=data['final_price']
    )
    combo.items.set(data['items'])
    combo.save()
    return JsonResponse({'message': '✅ Combo created successfully!'})


@require_POST
@login_required
@user_passes_test(is_admin)
def api_toggle_combo(request, combo_id):
    combo = get_object_or_404(Combo, pk=combo_id)
    combo.show_in_pos = not combo.show_in_pos
    combo.save()
    return JsonResponse({'status': combo.show_in_pos})

def reprint_receipt_view(request, order_id):
    order = get_object_or_404(Order, pk=order_id)
    return render(request, 'canteen/receipt.html', {'order': order})


#Report View
@login_required
def report(request):
    today = timezone.now().date()
    week_start = today - timedelta(days=today.weekday())  # Monday
    month_start = today.replace(day=1)
    year_start = today.replace(month=1, day=1)

    # Daily totals
    daily_sales = (
        Order.objects.filter(created_at__date=today)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )

    # Weekly totals
    weekly_sales = (
        Order.objects.filter(created_at__date__gte=week_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )

    # Monthly totals
    monthly_sales = (
        Order.objects.filter(created_at__date__gte=month_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )

    # ✅ Yearly totals (NEW)
    yearly_sales = (
        Order.objects.filter(created_at__date__gte=year_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )

    # Best-selling items
    top_items = (
        OrderItem.objects.values('product__name')
        .annotate(total_qty=Sum('quantity'))
        .order_by('-total_qty')[:5]
    )

    context = {
        "daily_sales": daily_sales,
        "weekly_sales": weekly_sales,
        "monthly_sales": monthly_sales,
        "yearly_sales": yearly_sales,  # ✅ pass to template
        "top_items": top_items,
    }

    return render(request, "canteen/report.html", context)


@login_required
def export_report_pdf(request):
    # Prepare data
    today = timezone.now().date()
    week_start = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)

    daily_sales = (
        Order.objects.filter(created_at__date=today)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    weekly_sales = (
        Order.objects.filter(created_at__date__gte=week_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    monthly_sales = (
        Order.objects.filter(created_at__date__gte=month_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    top_items = (
        OrderItem.objects.values('product__name')
        .annotate(total_qty=Sum('quantity'))
        .order_by('-total_qty')[:10]
    )

    # Create PDF
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    p.setTitle("Sales Report")

    p.setFont("Helvetica-Bold", 16)
    p.drawString(200, 810, "Canteen Sales Report")
    p.setFont("Helvetica", 12)
    p.drawString(50, 780, f"Date: {today}")
    p.drawString(50, 760, f"Daily Sales: ৳{daily_sales['total'] or 0}")
    p.drawString(50, 740, f"Weekly Sales: ৳{weekly_sales['total'] or 0}")
    p.drawString(50, 720, f"Monthly Sales: ৳{monthly_sales['total'] or 0}")

    y = 680
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "Top Selling Items")
    y -= 20
    p.setFont("Helvetica", 11)

    for item in top_items:
        p.drawString(60, y, f"{item['product__name']} - {item['total_qty']} sold")
        y -= 18
        if y < 50:
            p.showPage()
            y = 800

    p.showPage()
    p.save()

    pdf = buffer.getvalue()
    buffer.close()

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="canteen_report_{today}.pdf"'
    response.write(pdf)
    return response


@login_required
def export_report_excel(request):
    today = timezone.now().date()
    week_start = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)

    daily_sales = (
        Order.objects.filter(created_at__date=today)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    weekly_sales = (
        Order.objects.filter(created_at__date__gte=week_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    monthly_sales = (
        Order.objects.filter(created_at__date__gte=month_start)
        .aggregate(total=Sum('total_amount'), count=Count('id'))
    )
    top_items = (
        OrderItem.objects.values('product__name')
        .annotate(total_qty=Sum('quantity'))
        .order_by('-total_qty')[:10]
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Report"

    ws.append(["Canteen Sales Report"])
    ws.append(["Date", str(today)])
    ws.append([""])
    ws.append(["Daily Sales", daily_sales['total'] or 0])
    ws.append(["Weekly Sales", weekly_sales['total'] or 0])
    ws.append(["Monthly Sales", monthly_sales['total'] or 0])
    ws.append([""])
    ws.append(["Top Selling Items"])
    ws.append(["Item Name", "Quantity Sold"])

    for item in top_items:
        ws.append([item['product__name'], item['total_qty']])

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="canteen_report_{today}.xlsx"'

    with BytesIO() as buffer:
        wb.save(buffer)
        response.write(buffer.getvalue())

    return response