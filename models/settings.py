from odoo.exceptions import ValidationError
from odoo.osv import osv
from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools
from odoo.exceptions import UserError, AccessError

import requests
import json
import time

class Settings(models.Model):
    _inherit = 'res.users'

    field_name = fields.Char('Office365')

    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret_id = fields.Char('Secret Id')
    login_url = fields.Char('Login URL', compute='_compute_url', readonly=True)
    code = fields.Char('code')
    token = fields.Char('Token', readonly=True)
    refresh_token = fields.Char('Refresh Token', readonly=True)
    expires_in = fields.Char('Expires IN', readonly=True)
    office365_email = fields.Char('Office365 Email Address', readonly=True)
    office365_id_address = fields.Char('Office365 Id Address', readonly=True)
    from_date = fields.Datetime(string="From Date", required=False, )
    to_date = fields.Datetime(string="To Date", required=False, )



    def get_code(self):
        context = dict(self._context)
        settings = self.env['res.users'].search([])
        if self.redirect_url and self.client_id and self.login_url:
            if self.id == self.env.user.id:

                base_url = self.env['ir.config_parameter'].get_param('web.base.url')
                return {
                    'name': 'login',
                    'view_id': False,
                    "type": "ir.actions.act_url",
                    'target': 'self',
                    'url': self.login_url
                }
        else:
            raise ValidationError('Office365 Credentials are missing. Please! ask admin to add Office365 Client id, '
                                      'client secret and redirect Url ')

    @api.depends('redirect_url', 'client_id', 'secret_id')
    def _compute_url(self):

        # settings = self.env['res.users'].search([])
        # settings = settings[0] if settings else settings
        if self.redirect_url and self.client_id and self.secret_id:
            self.login_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' \
                             'client_id=%s&redirect_uri=%s&response_type=code&scope=openid+offline_access+' \
                             'Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+User.ReadWrite+Tasks.ReadWrite+' \
                             'Contacts.ReadWrite+MailboxSettings.Read' % (
                                 self.client_id, self.redirect_url)

    def test_connectiom(self):
        try:
            # settings = self.env['res.users'].search([])
            settings = self.env.user
            settings = settings[0] if settings else settings

            if not settings.client_id or not settings.redirect_url or not settings.secret_id:
                raise osv.except_osv(_("Error!"), (_("Please ask admin to add Office365 settings!")))

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data='grant_type=authorization_code&code=' + self.code + '&redirect_uri=' + settings.redirect_url + '&client_id=' + settings.client_id + '&client_secret=' + settings.secret_id
                , headers=header).content

            if 'error' in json.loads(response.decode('utf-8')) and json.loads(response.decode('utf-8'))['error']:
                raise UserError(
                    'Invalid Credentials . Please! Check your credential and  regenerate the code and try again!')

            else:
                response = json.loads((str(response)[2:])[:-1])
                self.env.user.token = response['access_token']
                self.env.user.refresh_token = response['refresh_token']

                self.env.user.expires_in = int(round(time.time() * 1000))
                self.env.user.code = self.code
                # self.code = ""
                response = json.loads((requests.get(
                    'https://graph.microsoft.com/v1.0/me',
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                        'Accept': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content.decode('utf-8')))
                self.env.user.office365_email = response['userPrincipalName']
                self.env.user.office365_id_address = 'outlook_' + response['id'].upper() + '@outlook.com'
                self.env.cr.commit()
        except Exception as e:
            raise ValidationError(_(str(e)))

    def generate_refresh_token(self):
        if self.env.user.refresh_token:
            settings = self.env.user
            # settings = self.env['res.users'].search([])
            settings = settings[0] if settings else settings

            if not settings.client_id or not settings.redirect_url or not settings.secret_id:
                raise osv.except_osv(_("Error!"), (_("Please ask admin to add Office365 settings!")))

            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }

            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data='grant_type=refresh_token&refresh_token=' + self.env.user.refresh_token + '&redirect_uri=' + settings.redirect_url + '&client_id=' + settings.client_id + '&client_secret=' + settings.secret_id
                , headers=header).content

            response = json.loads((str(response)[2:])[:-1])
            if 'access_token' not in response:
                response["error_description"] = response["error_description"].replace("\\r\\n", " ")
                raise osv.except_osv(("Error!"), (response["error"] + " " + response["error_description"]))
            else:
                self.env.user.token = response['access_token']
                self.env.user.refresh_token = response['refresh_token']
                self.env.user.expires_in = int(round(time.time() * 1000))
