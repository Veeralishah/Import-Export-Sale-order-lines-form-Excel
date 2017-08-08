# -*- coding: utf-8 -*-
{
    'name': "Import-Export_Sale_Order_Lines",

    'summary': """
        Import/Export Sale Order lines by Excel""",

    'description': """
        Import/Export Sale Order lines by Excel Sheet
    """,

    'author': "DRC Systems India Pvt. Ltd.",
    'website': "http://www.drcsystems.com/",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/master/odoo/addons/base/module/module_data.xml
    # for the full list
    'category': 'Sales',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': ['base','sale'],

    # always loaded
    'data': [
        'security/security.xml',
        'views/account_config_setting.xml',
        'wizard/sale_order.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
    ],
    'installble': True,
    'auto_install': False,
    'application': False,
}