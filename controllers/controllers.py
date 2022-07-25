# -*- coding: utf-8 -*-
# from odoo import http


# class Office365(http.Controller):
#     @http.route('/office365/office365', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/office365/office365/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('office365.listing', {
#             'root': '/office365/office365',
#             'objects': http.request.env['office365.office365'].search([]),
#         })

#     @http.route('/office365/office365/objects/<model("office365.office365"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('office365.object', {
#             'object': obj
#         })
