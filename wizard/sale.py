# -*- coding: utf-8 -*-

from odoo import models, fields, api, exceptions, _
import xlwt
import base64
from xlrd import open_workbook
from odoo.exceptions import UserError
try:
    from cStringIO import StringIO
except:
    from StringIO import StringIO



class Salewizard(models.TransientModel):
    _name = 'sale.exl'

    xl_data = fields.Binary(string='Select File')
    filename = fields.Char(string='filename')
    error_msg = fields.Text(string='Error')
    check_val = fields.Boolean(default=False)
    check_error = fields.Boolean(default=False)
    active_id = fields.Many2one(
        'sale.order', default=lambda self: self._context.get('active_id'))

    @api.multi
    def validate_file(self):
        text = ''
        if len(self._context.get('active_ids')) != 1:
            text = "import Sale order lines in only one record."
        else:
            if self.xl_data:
                if self.active_id.state != 'draft':
                    text = 'import order lines only in draft state.'
                else:
                    filename = str(self.filename)
                    if not (filename.endswith('xls') or filename.endswith('xlsx')):
                        text = "Import only '.xls' or '.xlsx' File."
                    else:
                        record_data = self.xl_data.decode('base64')
                        xls_open = open_workbook(file_contents=record_data)
                        column_list = ['ID', 'NAME', 'CODE',
                                       'QTY', 'PRICE', 'DISCOUNT']
                        value_column = []
                        for i in xls_open.sheets():
                            for col in range(i.ncols):
                                if not text:
                                    value = (i.cell(0, col).value)
                                    if value.upper() in value_column:
                                        text = "'" + value + "' columns is multiple time."
                                    else:
                                        value_column.append(value.upper())
                        if not text:
                            for j in value_column:
                                if not text:
                                    if j not in column_list:
                                        text = "'" + j + "' is a invalid Column name"
                            if 'ID' not in value_column and not text:
                                if 'CODE' not in value_column:
                                    text = "'Code' Column is compulsory"
            else:
                text = "File not selected"

        if not text:
            self.check_val = True
            return {
                'name': _(u'Import/Export Sale Order Lines'),
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'sale.exl',
                'target': 'new',
                'res_id': self.id,
                'context': self.env.context,
            }
        raise UserError(_(text))

    @api.multi
    def calculate_price(self, product, quantity):
        record_id = self.active_id
        product = product.with_context(
            lang=record_id.partner_id.lang,
            partner=record_id.partner_id.id,
            quantity=quantity,
            date=record_id.date_order,
            pricelist=record_id.pricelist_id.id,
        )
        if record_id.pricelist_id.discount_policy == 'without_discount':
            from_currency = record_id.company_id.currency_id
            price = from_currency.compute(product.lst_price, record_id.pricelist_id.currency_id)
        else:
            price = product.with_context(pricelist=record_id.pricelist_id.id).price
        return price

    @api.multi
    def order_import(self):
        record_data = self.xl_data.decode('base64')
        xls_open = open_workbook(file_contents=record_data)
        value_column = []
        default_qty = 0
        text = ''
        for i in xls_open.sheets():
            for row in range(i.nrows):
                data_row = []
                for col in range(i.ncols):
                    value = (i.cell(row, col).value)
                    if row == 0:
                        value_column.append(value.upper())
                    else:
                        data_row.append(value)
                vals = {}
                default_price = False
                if row != 0:
                    vals['price_unit'] = 0.0
                    if 'QTY' in value_column:
                        vals['product_uom_qty'] = data_row[
                            value_column.index('QTY')]
                    else:
                        vals['product_uom_qty'] = int(default_qty)
                    if 'NAME' in value_column:
                        vals['name'] = data_row[value_column.index('NAME')]
                    if 'PRICE' in value_column:
                        if data_row[value_column.index('PRICE')] != '':
                            vals['price_unit'] = data_row[
                                value_column.index('PRICE')]
                    if 'DISCOUNT' in value_column:
                        vals['discount'] = data_row[
                            value_column.index('DISCOUNT')]
                    if ('ID' not in value_column) or ('ID' in value_column and data_row[value_column.index('ID')] == ''):
                        if type(data_row[value_column.index('CODE')]) == float:
                            code = str(
                                int(data_row[value_column.index('CODE')]))
                        else:
                            code = str(data_row[value_column.index('CODE')])
                        product = self.env['product.product'].search(
                            [('default_code', '=', code)], limit=1)
                        if product:
                            vals[
                                'product_uom'] = product.product_tmpl_id.uom_id.id or False
                            if not vals['price_unit']:
                                vals['price_unit'] = self.calculate_price(
                                    product, vals['product_uom_qty'])
                                default_price = True
                            vals.update({'order_id': self.active_id.id,
                                         'product_id': product.id,
                                         'product_uom': product.product_tmpl_id.uom_id.id or False})
                            if not vals.get('name'):
                                vals[
                                    'name'] = product.product_tmpl_id.description_sale or product.name
                            new_line = self.env['sale.order.line'].create(vals)
                            if not default_price:
                                expected_subtotal = vals[
                                    'price_unit'] * int(vals['product_uom_qty'])
                                if not expected_subtotal == new_line.price_subtotal:
                                    price_difference = (
                                        expected_subtotal - new_line.price_subtotal) / int(vals['product_uom_qty'])
                                    new_line.write(
                                        {'price_unit': vals['price_unit'] + price_difference})
                        else:
                            text += "Product Code: " + str(data_row[value_column.index(
                                'CODE')]) + " invalid at Row " + str(row) +  "\n"
                    else:
                        order_line = self.env['sale.order.line'].search(
                            [('id', '=', int(data_row[0]))])
                        if order_line:
                            order_line.write(vals)
                            if not default_price:
                                new_line = order_line
                                expected_subtotal = vals[
                                    'price_unit'] * int(vals['product_uom_qty'])
                                if not expected_subtotal == new_line.price_subtotal:
                                    price_difference = (
                                        expected_subtotal - new_line.price_subtotal) / int(vals['product_uom_qty'])
                                    new_line.write(
                                        {'price_unit': vals['price_unit'] + price_difference})
                        else:
                            text += "ID: " + str(int(data_row[value_column.index(
                                'ID')])) + " invalid at Row " + str(row) + " and Column " + str(col) + "\n"
                else:
                    if 'QTY' not in value_column:
                        default_qty = self.env['ir.config_parameter'].search(
                            [('key', '=', 'Defauly Qty To Import Product in Sale Order')]).value
        if text:
            self.check_error = True
            self.error_msg = text
            return {
                'name': _(u'Import/Export Sale Order Lines'),
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'sale.exl',
                'target': 'new',
                'res_id': self.id,
                'context': self.env.context,
            }

    @api.multi
    def print_excel_file(self):
        self.ensure_one()

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
        sheet.write(0, 5, "Discount", cell)
        count_paid = 1

        for line in self.active_id.order_line:
            sheet.write(count_paid, 0, line.id or '')
            sheet.write(count_paid, 1, line.product_id.default_code or '')
            sheet.write(count_paid, 2, line.name or '')
            sheet.write(count_paid, 3, line.product_uom_qty or '')
            sheet.write(count_paid, 4, line.price_unit or '')
            sheet.write(count_paid, 5, line.discount or 0)
            count_paid += 1

        fp = StringIO()
        workbook.save(fp)
        fp.seek(0)
        data = fp.read()
        fp.close()
        return (base64.b64encode(data), sheet_name)

    @api.multi
    def order_export(self):
        if len(self._context.get('active_ids')) != 1:
            raise UserError(_("Export order lines only one record."))
        return {
                'type': 'ir.actions.act_url',
                'url': '/web/binary/download_document/SO/%s' % self.id,
                'target': 'self'
                }

class saleorder_wizard(models.Model):
    _inherit = 'sale.order'
