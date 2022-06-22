# -*- coding: utf-8 -*-

from odoo import models, fields, api


class office365(models.Model):
    _name = 'office365.office365'
    _description = 'office365.office365'

#     name = fields.Char()
#     value = fields.Integer()
#     value2 = fields.Float(compute="_value_pc", store=True)
#     description = fields.Text()
#
#     @api.depends('value')
#     def _value_pc(self):
#         for record in self:
#             record.value2 = float(record.value) / 100
