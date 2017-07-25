# -*- coding : utf-8 -*-

from odoo import models, fields, api
from odoo.osv import osv


class SaleConfiguration(models.TransientModel):
	_name = 'sale.config.settings'
	_inherit = 'sale.config.settings'

	group_opt = fields.Boolean(string = "Import Sale Order line From Excel file?", defaults = '',
	implied_group='import-export_sale_order.group_sale_impex_orderline')