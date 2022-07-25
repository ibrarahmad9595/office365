# -*- coding: utf-8 -*-
{
    'name': "office365",

    'summary': """
        Email can be send from odoo to office365 and other email service. 
        Email can be received in odoo from office365 and other email service. """,

    'description': """
        Odoo is a suite of business management software tools including, for example,
        CRM, e-commerce, billing, accounting, manufacturing, warehouse, project management,
        and inventory management.
    """,

    'author': "Solocrew",
    'website': "http://www.solocrew.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/15.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '15.0.0.1',
    'sequence':'-100',
    # any module necessary for this one to work correctly
    'depends': ['base', 'calendar', 'crm'],
    'images': [
        'static/description/banner.png',
    ],

    # always loaded
    'data': [
        'security/ir.model.access.csv',
        'views/views.xml',
        'views/templates.xml',
    ],
    # only loaded in demonstration mode
    'demo': [
        'demo/demo.xml',
    ],
}
