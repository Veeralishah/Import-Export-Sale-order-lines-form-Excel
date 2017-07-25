# -*- coding: utf-8 -*-
import odoo
from odoo.tests.common import TransactionCase
import xlrd
import xlwt
import os
try:
    from cStringIO import StringIO
except:
    from StringIO import StringIO
@odoo.tests.common.post_install(True)
@odoo.tests.common.at_install(False)

class TestImpexSO(TransactionCase):


    def test_SO(self):
        SO = self.env['sale.order']
        SO_line = self.env['sale.order.line']
        user = self.env['res.partner']
        product = self.env['product.product']
        pricelist = self.env['product.pricelist'].create({'name': 'test_pricelist'})
        partner = user.create({'name': 'test'})
        so = SO.create({'partner_id': partner.id, 'pricelist_id': pricelist.id, 'picking_policy': 'direct'})
        path = os.path.dirname(os.path.abspath(__file__)) + '/test_so.xlsx'
        workbook = xlrd.open_workbook(path)
        for s in workbook.sheets():
            for row in range(s.nrows):
                prod_code = s.cell(row, 0)
                prod_name = s.cell(row, 1)
                prod_qty = s.cell(row, 2).value
                prod_price = s.cell(row, 3).value
                product = product.create({'default_code': prod_code, 'name': prod_name})
                SO_line = SO_line.create({'order_id': so.id,
                            'product_id': product.id,
                            'name': product.name,
                            'product_uom': product.product_tmpl_id.uom_po_id.id,
                            'price_unit': prod_price,
                            'product_uom_qty': prod_qty
                            })

        sheet_name = 'Sale Order Line'
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Sale Order Line 1')
        sheet.col(0).width = 256 * 35
        sheet.col(1).width = 256 * 20
        sheet.col(2).width = 256 * 20
        sheet.col(3).width = 256 * 40
        sheet.col(4).width = 256 * 20
        sheet.col(5).width = 256 * 20

        font = xlwt.Font()
        style = xlwt.XFStyle()
        style.font = font
        heading = xlwt.easyxf('font: bold on, height 300; align: horiz center;')
        bold = xlwt.easyxf('font: bold on')
        cell = xlwt.easyxf('font: bold on, height 200; align: horiz center;')
        total = xlwt.easyxf('font: bold on, height 220; align: horiz right;')
        center = xlwt.easyxf('align: horiz center;')

        sheet.write(0, 0, "ID", cell)
        sheet.write(0, 1, "Code", cell)
        sheet.write(0, 2, "Name", cell)
        sheet.write(0, 3, "Qty", cell)
        sheet.write(0, 4, "Price", cell)
        count_paid = 1

        for line in so.order_line:
            sheet.write(count_paid, 0, line.id or '')
            sheet.write(count_paid, 1, line.product_id.default_code or '')
            sheet.write(count_paid, 2, line.name or '')
            sheet.write(count_paid, 3, line.product_uom_qty or '')
            sheet.write(count_paid, 4, line.price_unit or '')
            count_paid += 1
        fp = StringIO()
        workbook.save(fp)
        fp.seek(0)
        data = fp.read()
        fp.close()