from django.test import TestCase
from django.template.loader import get_template

from posApp.models import Category, Products, Sales, salesItems


class ReceiptTemplateTests(TestCase):
    def test_receipt_shows_qty_for_each_product(self):
        category = Category.objects.create(name='Test Category', description='Test')
        product = Products.objects.create(
            code='P001',
            category_id=category,
            name='Test Product',
            price=100,
            modal=50,
            stock=10,
        )
        sale = Sales.objects.create(
            code='TEST001',
            sub_total=100,
            grand_total=100,
            tendered_amount=100,
            amount_change=0,
            profit=50,
        )
        salesItems.objects.create(
            sale_id=sale,
            product_id=product,
            price=100,
            modal=50,
            qty=3,
            total=300,
            profit=150,
        )

        template = get_template('posApp/receipt.html')
        rendered = template.render({
            'transaction': {
                'date_added': sale.date_added,
                'code': sale.code,
                'grand_total': sale.grand_total,
                'tendered_amount': sale.tendered_amount,
                'amount_change': sale.amount_change,
            },
            'salesItems': salesItems.objects.filter(sale_id=sale),
        })

        self.assertIn('3', rendered)
