from pickle import FALSE
from django.shortcuts import redirect, render
from django.http import HttpResponse
from flask import jsonify
from posApp.models import Category, Products, Sales, salesItems, Restock
from django.utils.safestring import mark_safe
from django.db.models import Count, Sum
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect
import json, sys
from datetime import date, datetime
from django.db.models import OuterRef, Subquery
import traceback
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from django.http import HttpResponse

# Login
def login_user(request):
    logout(request)
    resp = {"status":'failed','msg':''}
    username = ''
    password = ''
    if request.POST:
        username = request.POST['username']
        password = request.POST['password']

        user = authenticate(username=username, password=password)
        if user is not None:
            if user.is_active:
                login(request, user)
                resp['status']='success'
            else:
                resp['msg'] = "Username atau Password salah!"
        else:
            resp['msg'] = "Username atau Password salah!"
    return HttpResponse(json.dumps(resp),content_type='application/json')

#Logout
def logoutuser(request):
    logout(request)
    return redirect('/')

# Create your views here.
@login_required
def home(request):
    now = datetime.now()
    current_year = now.strftime("%Y")
    current_month = now.strftime("%m")
    current_day = now.strftime("%d")
    categories = len(Category.objects.all())
    products = len(Products.objects.all())
    transaction = len(Sales.objects.filter(
        date_added__year=current_year,
        date_added__month = current_month,
        date_added__day = current_day
    ))
    today_sales = Sales.objects.filter(
        date_added__year=current_year,
        date_added__month = current_month,
        date_added__day = current_day
    ).all()
    total_sales = sum(today_sales.values_list('grand_total',flat=True))
    context = {
        'page_title':'Home',
        'categories' : categories,
        'products' : products,
        'transaction' : transaction,
        'total_sales' : total_sales,
    }
    return render(request, 'posApp/home.html',context)


def about(request):
    context = {
        'page_title':'About',
    }
    return render(request, 'posApp/about.html',context)

#Categories
@login_required
def category(request):
    category_list = Category.objects.all()
    context = {
        'page_title':'Category List',
        'category':category_list,
    }
    return render(request, 'posApp/category.html',context)
@login_required
def manage_category(request):
    category = {}
    if request.method == 'GET':
        data =  request.GET
        id = ''
        if 'id' in data:
            id= data['id']
        if id.isnumeric() and int(id) > 0:
            category = Category.objects.filter(id=id).first()
    
    context = {
        'category' : category
    }
    return render(request, 'posApp/manage_category.html',context)

@login_required
def save_category(request):
    data =  request.POST
    resp = {'status':'failed'}
    try:
        if (data['id']).isnumeric() and int(data['id']) > 0 :
            save_category = Category.objects.filter(id = data['id']).update(name=data['name'], description = data['description'])
        else:
            save_category = Category(name=data['name'], description = data['description'])
            save_category.save()
        resp['status'] = 'success'
        messages.success(request, 'Kategori berhasil disimpan.')
    except:
        resp['status'] = 'failed'
    return HttpResponse(json.dumps(resp), content_type="application/json")

@login_required
def delete_category(request):
    data =  request.POST
    resp = {'status':''}
    try:
        Category.objects.filter(id = data['id']).delete()
        resp['status'] = 'success'
        messages.success(request, 'Category Successfully deleted.')
    except:
        resp['status'] = 'failed'
    return HttpResponse(json.dumps(resp), content_type="application/json")

# Products
@login_required
def products(request):
    latest_restock = Restock.objects.filter(
        product_id=OuterRef('pk')
    ).order_by('-restock_date')

    product_list = Products.objects.annotate(
        latest_f_restock=Subquery(latest_restock.values('f_restock')[:1]),
        latest_sale=Subquery(latest_restock.values('sale')[:1]),
        latest_restock_date=Subquery(latest_restock.values('restock_date')[:1])
    )
    context = {
        'page_title':'Product List',
        'products':product_list,
    }
    return render(request, 'posApp/products.html',context)

def restock_list(request):
    # Ambil data restock terakhir per produk
    latest_restock = Restock.objects.filter(
        product_id=OuterRef('pk')
    ).order_by('-restock_date')

    # Ambil produk dengan stok kurang dari 5
    restock_products = Products.objects.filter(stock__lt=5).annotate(
        latest_f_restock=Subquery(latest_restock.values('f_restock')[:1]),
        latest_sale=Subquery(latest_restock.values('sale')[:1]),
        latest_restock_date=Subquery(latest_restock.values('restock_date')[:1])
    ).order_by('stock')

    context = {
        'page_title': 'Daftar produk yang perlu di restok',
        'products': restock_products,
    }

    return render(request, 'posApp/restock_list.html', context)

@login_required
def manage_products(request):
    product = {}
    categories = Category.objects.all()

    if request.method == 'GET':
        data = request.GET
        id = data.get('id', '')

        if id.isnumeric() and int(id) > 0:
            product = Products.objects.filter(id=id).first()

    # --- Dapatkan kode terakhir ---
    last_product = Products.objects.order_by('-id').first()
    if last_product:
        # Ambil angka dari kode terakhir (contoh: 005 -> 5)
        last_code_num = int(''.join(filter(str.isdigit, last_product.code)))
        next_code_num = last_code_num + 1
    else:
        next_code_num = 1

    # Format kode baru misal: P001, P002, ...
    next_code = f"{next_code_num:03d}"

    context = {
        'product': product,
        'categories': categories,
        'next_code': next_code
    }
    return render(request, 'posApp/manage_product.html', context)

def test(request):
    categories = Category.objects.all()
    context = {
        'categories' : categories
    }
    return render(request, 'posApp/test.html',context)

@login_required
def save_product(request):
    data =  request.POST
    resp = {'status':'failed'}
    id= ''
    if 'id' in data:
        id = data['id']
    if id.isnumeric() and int(id) > 0:
        check = Products.objects.exclude(id=id).filter(code=data['code']).all()
    else:
        check = Products.objects.filter(code=data['code']).all()
    if len(check) > 0 :
        resp['msg'] = "Product Code Already Exists in the database"
    else:
        category = Category.objects.filter(id = data['category_id']).first()
        
        try:
            # Update
            if (data['id']).isnumeric() and int(data['id']) > 0 :
                add_stock = int(data['stock'])
                new_stock = Products.objects.filter(id = data['id']).first().stock + add_stock
                save_product = Products.objects.filter(id = data['id']).update(code=data['code'], category_id=category, name=data['name'], price = float(data['price']), stock=new_stock)
                restock = Restock.objects.filter(product_id=data['id']).last()
                Restock.objects.filter(id=restock.pk).update(sale=0, f_restock_before=restock.f_restock, f_restock=0)
            # create   
            else:
                stock = int(data['stock'])
                status = True if stock > 0 else False
                save_product = Products(code=data['code'], category_id=category, name=data['name'], price = float(data['price']), status = status, stock=int(data['stock']))
                save_product.save()
                Restock(product_id=save_product, f_restock_before=stock).save()

            resp['status'] = 'success'
            messages.success(request, 'Product Successfully saved.')
        except:
            resp['status'] = 'failed'
    return HttpResponse(json.dumps(resp), content_type="application/json")

@login_required
def delete_product(request):
    data =  request.POST
    resp = {'status':''}
    try:
        Products.objects.filter(id = data['id']).delete()
        resp['status'] = 'success'
        messages.success(request, 'Product Successfully deleted.')
    except:
        resp['status'] = 'failed'
    return HttpResponse(json.dumps(resp), content_type="application/json")

@login_required
def pos(request):
    products = Products.objects.filter(status = 1)
    product_json = []
    for product in products:
        product_json.append({'id':product.id, 'name':product.name, 'price':float(product.price), 'stock':int(product.stock)})
    context = {
        'page_title' : "Point of Sale",
        'products' : products,
        'product_json' : json.dumps(product_json)
    }
    # return HttpResponse('')
    return render(request, 'posApp/pos.html',context)

@login_required
def checkout_modal(request):
    grand_total = 0
    if 'grand_total' in request.GET:
        grand_total = request.GET['grand_total']
    context = {
        'grand_total' : grand_total,
    }
    return render(request, 'posApp/checkout.html', context)

def predict_f_restock(id):
    old_restock = Restock.objects.filter(pk=id).first()
    new_f_restock = (0.1 * old_restock.sale) + (1 - 0.1) * old_restock.f_restock_before
    Restock.objects.filter(pk=id).update(f_restock=new_f_restock)

@login_required
def save_pos(request):
    resp = {'status':'failed','msg':''}
    data = request.POST
    pref = datetime.now().year + datetime.now().year
    i = 1
    while True:
        code = '{:0>5}'.format(i)
        i += int(1)
        check = Sales.objects.filter(code = str(pref) + str(code)).all()
        if len(check) <= 0:
            break
    code = str(pref) + str(code)

    try:
        sales = Sales(code=code, sub_total = data['sub_total'], grand_total = data['grand_total'], tendered_amount = data['tendered_amount'], amount_change = data['amount_change']).save()
        sale_id = Sales.objects.last().pk
        i = 0
        for prod in data.getlist('product_id[]'):
            product_id = prod 
            sale = Sales.objects.filter(id=sale_id).first()
            product = Products.objects.filter(id=product_id).first()
            latest_restock = Restock.objects.filter(product_id=product_id).last()
            qty = data.getlist('qty[]')[i]
            terjual = latest_restock.sale 
            price = data.getlist('price[]')[i] 
            total = float(qty) * float(price)
            stock_now = product.stock - int(qty)
            sale_now = terjual + int(qty)
            Products.objects.filter(id=product_id).update(stock=stock_now)
            Restock.objects.filter(id=latest_restock.pk).update(sale=sale_now)
            predict_f_restock(latest_restock.pk)
            print({'sale_id' : sale, 'product_id' : product, 'qty' : qty, 'price' : price, 'total' : total})
            salesItems(sale_id = sale, product_id = product, qty = qty, price = price, total = total).save()
            i += int(1)
            
        resp['status'] = 'success'
        resp['sale_id'] = sale_id
        messages.success(request, "Sale Record has been saved.")

        # 🟦 Tambahan: cek stok rendah dan tampilkan alert
        low_stock_products = Products.objects.filter(stock__lt=5)
        if low_stock_products.exists():
            low_list = ", ".join([p.name for p in low_stock_products])
            alert_message = (
                f"⚠️ Stok menipis untuk produk: {low_list}.<br>"
                f"<a href='/restock-list/' class='btn btn-sm btn-primary mt-2'>"
                f"Lihat Daftar Restok</a>"
            )
            messages.warning(request, mark_safe(alert_message))  # ✅ aman untuk HTML dalam messages

    except Exception as e:
        resp['msg'] = "An error occured"
        print("Unexpected error:", sys.exc_info()[0])
        resp['msg'] = f"An error occurred: {str(e)}"
        print("Unexpected error:", e)
        traceback.print_exc()
    return HttpResponse(json.dumps(resp),content_type="application/json")


@login_required
def salesList(request):
    sales = Sales.objects.all().order_by('-date_added')
    month = request.GET.get('month', '')
    start_month = request.GET.get('start_month', '')
    end_month = request.GET.get('end_month', '')
    year = request.GET.get('year', '')

    if year:
        sales = sales.filter(date_added__year=year)
    if month:
        sales = sales.filter(date_added__month=month)
    elif start_month and end_month:
        sales = sales.filter(date_added__month__gte=start_month, date_added__month__lte=end_month)

    sale_data = []
    for sale in sales:
        data = {}
        for field in sale._meta.get_fields(include_parents=False):
            if field.related_model is None:
                data[field.name] = getattr(sale,field.name)
        data['items'] = salesItems.objects.filter(sale_id=sale).all()
        data['item_count'] = len(data['items'])
        if data['item_count'] <= 0:
            continue
        if 'tax_amount' in data:
            data['tax_amount'] = format(float(data['tax_amount']),'.2f')
        sale_data.append(data)

    year_choices = list(Sales.objects.dates('date_added', 'year', order='DESC'))
    year_choices = [y.year for y in year_choices]
    month_choices = [
        (1, 'Januari'), (2, 'Februari'), (3, 'Maret'), (4, 'April'),
        (5, 'Mei'), (6, 'Juni'), (7, 'Juli'), (8, 'Agustus'),
        (9, 'September'), (10, 'Oktober'), (11, 'November'), (12, 'Desember')
    ]

    context = {
        'page_title':'Sales Transactions',
        'sale_data':sale_data,
        'selected_month': month,
        'selected_start_month': start_month,
        'selected_end_month': end_month,
        'selected_year': year,
        'year_choices': year_choices,
        'month_choices': month_choices,
    }
    return render(request, 'posApp/sales.html', context)

@login_required
def export_sales_excel(request):
    month = request.GET.get('month', '')
    start_month = request.GET.get('start_month', '')
    end_month = request.GET.get('end_month', '')
    year = request.GET.get('year', '')

    sales = Sales.objects.all().order_by('-date_added')
    if year:
        sales = sales.filter(date_added__year=year)
    if month:
        sales = sales.filter(date_added__month=month)
    elif start_month and end_month:
        sales = sales.filter(date_added__month__gte=start_month, date_added__month__lte=end_month)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Data Transaksi'

    headers = ['Tanggal', 'Kode Transaksi', 'Subtotal', 'Grand Total', 'Dibayar', 'Kembalian']
    ws.append(headers)

    for sale in sales:
        ws.append([
            sale.date_added.strftime('%Y-%m-%d %H:%M'),
            sale.code,
            sale.sub_total,
            sale.grand_total,
            sale.tendered_amount,
            sale.amount_change,
        ])

    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill("solid", fgColor="4F81BD")
        cell.alignment = Alignment(horizontal='center')

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="laporan_transaksi.xlsx"'
    wb.save(response)
    return response

@login_required
def receipt(request):
    id = request.GET.get('id')
    sales = Sales.objects.filter(id = id).first()
    transaction = {}
    for field in Sales._meta.get_fields():
        if field.related_model is None:
            transaction[field.name] = getattr(sales,field.name)
    if 'tax_amount' in transaction:
        transaction['tax_amount'] = format(float(transaction['tax_amount']))
    ItemList = salesItems.objects.filter(sale_id = sales).all()
    context = {
        "transaction" : transaction,
        "salesItems" : ItemList
    }

    return render(request, 'posApp/receipt.html',context)
    # return HttpResponse('')

@login_required
def delete_sale(request):
    resp = {'status':'failed', 'msg':''}
    id = request.POST.get('id')
    try:
        delete = Sales.objects.filter(id = id).delete()
        resp['status'] = 'success'
        messages.success(request, 'Sale Record has been deleted.')
    except:
        resp['msg'] = "An error occured"
        print("Unexpected error:", sys.exc_info()[0])
    return HttpResponse(json.dumps(resp), content_type='application/json')

def guest_product_display(request):
    # Ambil semua produk
    products = Products.objects.all().order_by('name')
    context = {
        'page_title': 'Daftar Produk Koperasi',
        'products': products
    }
    return render(request, 'posApp/guest_display.html', context)