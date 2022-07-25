import logging
_logger = logging.getLogger(__name__)

from odoo.exceptions import ValidationError
from odoo.osv import osv
from odoo import _, api, fields, models, modules, SUPERUSER_ID, tools
from odoo.exceptions import UserError, AccessError
import requests
import json
from datetime import datetime

import time
from datetime import timedelta


class Office365Configuration(models.Model):
    _name = 'office365.configuration'
    field_name = fields.Char('Office365')

    login_url = fields.Char('Login URL', compute='_compute_url', readonly=True)
    code = fields.Char('code')
    redirect_url = fields.Char('Redirect URL')
    client_id = fields.Char('Client Id')
    secret_id = fields.Char('Secret Id')

    def save(self):
        self.env.user.redirect_url = self.redirect_url;
        self.env.user.client_id = self.client_id;
        self.env.user.secret_id = self.secret_id;
        self.env.cr.commit()

    def save_code(self):
        self.env.user.code = self.code;
        self.env.cr.commit()

    def get_code(self):
        context = dict(self._context)
        settings = self.env['res.users'].search([])
        if self.env.user.redirect_url and self.env.user.client_id and self.env.user.login_url:
            # if self.id == self.env.user.id:
                base_url = self.env['ir.config_parameter'].get_param('web.base.url')
                return {
                    'name': 'login',
                    'view_id': False,
                    "type": "ir.actions.act_url",
                    'target': 'self',
                    'url': self.env.user.login_url
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

class Office365Integration(models.Model):
    _name = 'office365.integration'
    field_name = fields.Char('Office365')
    import_contact = fields.Boolean('Import contacts', strore=True)
    export_contact = fields.Boolean(string='Export contacts')
    import_email = fields.Boolean(string='Import contacts email ')

    from_date = fields.Datetime(string="From Date", required=False, )
    to_date = fields.Datetime(string="To Date", required=False, )

    def sync_data(self):
        if self.import_contact:
            self.import_contacts()
        if self.export_contact:
            self.export_contacts()
        if self.import_email:
            self.sync_customer_mail()

    def import_contacts(self):
        """
        This is for importing contacts to office 365
        :return:
        """
        office_contacts = []
        if self.env.user.token:
            try:
                if self.env.user.token:
                    if self.env.user.expires_in:
                        expires_in = datetime.fromtimestamp(int(self.env.user.expires_in) / 1e3)
                        expires_in = expires_in + timedelta(seconds=3600)
                        nowDateTime = datetime.now()
                        if nowDateTime > expires_in:
                            self.generate_refresh_token()
                    count = 0
                    test_count = 0
                    url = 'https://graph.microsoft.com/v1.0/me/contacts?$top=500'
                    headers = {
                        'Host': 'outlook.office365.com',
                        'Authorization': 'Bearer {0}'.format(self.env.user.token),
                        'Accept': 'application/json',
                        'Con#'
                        'tent-Type': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }
                    while True:
                        response = requests.get(
                            url, headers=headers
                        )
                        if response.status_code == 404:
                            break
                        response = json.loads(response.content.decode('utf-8'))
                        # response = json.loads(response.decode('utf-8'))
                        if not ('value' in response and response['value']):
                            break

                        phone = None
                        print('1')
                        if 'value' in response:
                            for each_contact in response['value']:
                                print(2)

                                if ('emailAddresses' not in each_contact or not each_contact['displayName']) and (
                                        'emailAddresses' not in each_contact or not each_contact[
                                    'emailAddresses']) and (
                                        not 'businessPhones' in each_contact or not each_contact['businessPhones']):
                                    continue
                                odoo_cust = self.env['res.partner'].search(
                                    [('office_contact_id', '=', each_contact['id'])])
                                if not odoo_cust:
                                    print('3')
                                    if ('emailAddresses' in each_contact and len(
                                            each_contact['emailAddresses']) > 0 and (
                                                'emailAddresses' in each_contact and each_contact['emailAddresses'][0][
                                            'address'] != None)) or ((
                                            'mobilePhone' in each_contact and each_contact['mobilePhone'] or len(
                                            'homePhones' in each_contact and each_contact['homePhones']) > 0 or len(
                                            'businessPhones' in each_contact and each_contact['businessPhones']) > 0)):
                                        print('4')
                                        if each_contact['emailAddresses'] and each_contact['emailAddresses'][0][
                                            'address']:
                                            email_address = each_contact['emailAddresses'][0]['address']
                                            if office_contacts:
                                                office_contact_id = [i for i in office_contacts if
                                                                     i['email'] == email_address]
                                                if office_contact_id:
                                                    print(office_contact_id)
                                                    continue
                                            print('4.1')

                                        if each_contact['homePhones'] and len(each_contact['homePhones']) > 0:
                                            phone = each_contact['homePhones'][0]
                                            if office_contacts:
                                                office_contact_id = [i for i in office_contacts if
                                                                     i['phone'] == phone]
                                                if office_contact_id:
                                                    print(office_contact_id)
                                                    continue
                                        elif each_contact['businessPhones'] and len(each_contact['businessPhones']) > 0:
                                            phone = each_contact['businessPhones'][0]
                                            if office_contacts:
                                                office_contact_id = [i for i in office_contacts if
                                                                     i['phone'] == phone]
                                                if office_contact_id:
                                                    print(office_contact_id)

                                        if phone and email_address:
                                            if not self.env['res.partner'].search(
                                                    ['|', ('email', '=', email_address), ('phone', '=', phone)]):
                                                contact_data = {
                                                    'company_id': self.env.user.company_id.id,
                                                    'name': each_contact[
                                                        'displayName'] if 'displayName' in each_contact else
                                                    email_address,
                                                    'email': email_address if email_address else None,
                                                    'company_name': each_contact[
                                                        'companyName'] if 'companyName' in each_contact else None,
                                                    'function': each_contact[
                                                        'jobTitle'] if 'jobTitle' in each_contact else None,
                                                    'office_contact_id': each_contact['id'],
                                                    'mobile': each_contact[
                                                        'mobilePhone'] if 'mobilePhone' in each_contact else None,
                                                    'phone': phone if phone else None,
                                                    'street': each_contact['homeAddress']['street'] if each_contact[
                                                        'homeAddress'] else None,
                                                    'city': each_contact['homeAddress']['city'] if 'city' in
                                                                                                   each_contact[
                                                                                                       'homeAddress'] and
                                                                                                   each_contact[
                                                                                                       'homeAddress'] else None,
                                                    'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in
                                                                                                        each_contact[
                                                                                                            'homeAddress'] and
                                                                                                        each_contact[
                                                                                                            'homeAddress'] else None,
                                                    'state_id': self.env['res.country.state'].search(
                                                        [('name', '=', each_contact['homeAddress']['state'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                    'country_id': self.env['res.country'].search(
                                                        [('name', '=',
                                                          each_contact['homeAddress']['countryOrRegion'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                }

                                                office_contacts.append(contact_data)

                                                self.env.cr.commit()
                                        elif email_address:
                                            if not self.env['res.partner'].search(
                                                    [('email', '=', email_address)]):
                                                contact_data = {
                                                    'company_id': self.env.user.company_id.id,
                                                    'name': each_contact[
                                                        'displayName'] if 'displayName' in each_contact else
                                                    email_address,
                                                    'email': email_address if email_address else None,
                                                    'company_name': each_contact[
                                                        'companyName'] if 'companyName' in each_contact else None,
                                                    'function': each_contact[
                                                        'jobTitle'] if 'jobTitle' in each_contact else None,
                                                    'office_contact_id': each_contact['id'],
                                                    'mobile': each_contact[
                                                        'mobilePhone'] if 'mobilePhone' in each_contact else None,
                                                    'phone': phone if phone else None,
                                                    'street': each_contact['homeAddress']['street'] if each_contact[
                                                        'homeAddress'] else None,
                                                    'city': each_contact['homeAddress']['city'] if 'city' in
                                                                                                   each_contact[
                                                                                                       'homeAddress'] and
                                                                                                   each_contact[
                                                                                                       'homeAddress'] else None,
                                                    'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in
                                                                                                        each_contact[
                                                                                                            'homeAddress'] and
                                                                                                        each_contact[
                                                                                                            'homeAddress'] else None,
                                                    'state_id': self.env['res.country.state'].search(
                                                        [('name', '=', each_contact['homeAddress']['state'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                    'country_id': self.env['res.country'].search(
                                                        [('name', '=',
                                                          each_contact['homeAddress']['countryOrRegion'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                }

                                                office_contacts.append(contact_data)

                                                self.env.cr.commit()
                                        elif phone:
                                            if not self.env['res.partner'].search(
                                                    [('phone', '=', phone)]):
                                                contact_data = {
                                                    'company_id': self.env.user.company_id.id,
                                                    'name': each_contact[
                                                        'displayName'] if 'displayName' in each_contact else
                                                    email_address,
                                                    'email': email_address if email_address else None,
                                                    'company_name': each_contact[
                                                        'companyName'] if 'companyName' in each_contact else None,
                                                    'function': each_contact[
                                                        'jobTitle'] if 'jobTitle' in each_contact else None,
                                                    'office_contact_id': each_contact['id'],
                                                    'mobile': each_contact[
                                                        'mobilePhone'] if 'mobilePhone' in each_contact else None,
                                                    'phone': phone if phone else None,
                                                    'street': each_contact['homeAddress']['street'] if each_contact[
                                                        'homeAddress'] else None,
                                                    'city': each_contact['homeAddress']['city'] if 'city' in
                                                                                                   each_contact[
                                                                                                       'homeAddress'] and
                                                                                                   each_contact[
                                                                                                       'homeAddress'] else None,
                                                    'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in
                                                                                                        each_contact[
                                                                                                            'homeAddress'] and
                                                                                                        each_contact[
                                                                                                            'homeAddress'] else None,
                                                    'state_id': self.env['res.country.state'].search(
                                                        [('name', '=', each_contact['homeAddress']['state'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                    'country_id': self.env['res.country'].search(
                                                        [('name', '=',
                                                          each_contact['homeAddress']['countryOrRegion'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                }

                                                office_contacts.append(contact_data)

                                                self.env.cr.commit()
                                        elif 'mobilePhone' in each_contact and each_contact['mobilePhone']:
                                            if not self.env['res.partner'].search(
                                                    ['|', ('email', '=', email_address), ('phone', '=', phone)]):
                                                contact_data = {
                                                    'company_id': self.env.user.company_id.id,
                                                    'name': each_contact[
                                                        'displayName'] if 'displayName' in each_contact else
                                                    email_address,
                                                    'email': email_address if email_address else None,
                                                    'company_name': each_contact[
                                                        'companyName'] if 'companyName' in each_contact else None,
                                                    'function': each_contact[
                                                        'jobTitle'] if 'jobTitle' in each_contact else None,
                                                    'office_contact_id': each_contact['id'],
                                                    'mobile': each_contact[
                                                        'mobilePhone'] if 'mobilePhone' in each_contact else None,
                                                    'phone': phone if phone else None,
                                                    'street': each_contact['homeAddress']['street'] if each_contact[
                                                        'homeAddress'] else None,
                                                    'city': each_contact['homeAddress']['city'] if 'city' in
                                                                                                   each_contact[
                                                                                                       'homeAddress'] and
                                                                                                   each_contact[
                                                                                                       'homeAddress'] else None,
                                                    'zip': each_contact['homeAddress']['postalCode'] if 'postalCode' in
                                                                                                        each_contact[
                                                                                                            'homeAddress'] and
                                                                                                        each_contact[
                                                                                                            'homeAddress'] else None,
                                                    'state_id': self.env['res.country.state'].search(
                                                        [('name', '=', each_contact['homeAddress']['state'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                    'country_id': self.env['res.country'].search(
                                                        [('name', '=',
                                                          each_contact['homeAddress']['countryOrRegion'])]).id if
                                                    each_contact['homeAddress'] else None,
                                                }
                                                office_contacts.append(contact_data)

                                                self.env.cr.commit()
                            print('5')
                            if '@odata.nextLink' in response:
                                print('6')
                                url = response['@odata.nextLink']
                                print('7')

                            else:
                                break
                    if office_contacts:
                        odoo_contact = self.env['res.partner'].create(office_contacts)
                        self.env.cr.commit()
                        # print(odoo_contact)

                else:
                    raise UserWarning('Token is missing. Please Generate Token ')

            except Exception as e:
                raise ValidationError(_(str(e)))

    def generate_refresh_token(self):

        if self.env.user.refresh_token:
            settings = self.env['res.users'].search([])
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

    def sync_customer_mail(self):
        try:
            context = self._context
            current_uid = context.get('uid')
            res_user = self.env['res.users'].browse(current_uid)
            self.env.cr.commit()
            self.sync_customer_inbox_mail()
            self.sync_customer_sent_mail()

        except Exception as e:
            self.env.cr.commit()
            raise ValidationError(_(str(e)))
        self.env.cr.commit()

    def sync_customer_inbox_mail(self):
        context = self._context
        current_uid = context.get('uid')
        res_user = self.env['res.users'].browse(current_uid)
        new_email = []
        status = None
        if res_user.token:
            try:
                if res_user.expires_in:
                    expires_in = datetime.fromtimestamp(int(res_user.expires_in) / 1e3)
                    expires_in = expires_in + timedelta(seconds=3600)
                    nowDateTime = datetime.now()
                    if nowDateTime > expires_in:
                        self.generate_refresh_token()
                    url = 'https://graph.microsoft.com/v1.0/me/mailFolders?$top=500'
                    headers = {
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(res_user.token),
                        'Accept': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'

                    }

                response = requests.get(url,
                    headers=headers)

                if 'value' not in json.loads((response.content.decode('utf-8'))).keys():
                    raise osv.except_osv("Access TOken Expired!", " Please Regenerate Access Token !")
                folders = json.loads((response.content.decode('utf-8')))['value']
                inbox_id = [folder['id'] for folder in folders if folder['displayName'] == 'กล่องจดหมายเข้า' or folder['displayName'] == 'Inbox'] #'Inbox' or
                if inbox_id:
                    inbox_id = inbox_id[0]
                    url = ''
                    custom_data = self.env['office365.integration'].search([])[0]
                    if custom_data.from_date and custom_data.to_date:
                        url = 'https://graph.microsoft.com/v1.0/me/mailFolders/' + inbox_id + \
                              '/messages?$top=1000&$count=true&$filter=ReceivedDateTime ge {} & ReceivedDateTime le {}' \
                                  .format(custom_data.from_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
                                          custom_data.to_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
                    else:
                        url = 'https://graph.microsoft.com/v1.0/me/mailFolders/' + inbox_id + '/messages'
                        while True:
                            if response.status_code == 404:
                                break
                            response = requests.get(url,
                                                    headers={
                                                        'Host': 'outlook.office.com',
                                                        'Authorization': 'Bearer {0}'.format(res_user.token),
                                                        'Accept': 'application/json',
                                                        'X-Target-URL': 'http://outlook.office.com',
                                                        'connection': 'keep-Alive'
                                                    })

                            if 'value' not in json.loads((response.content.decode('utf-8'))).keys():
                                raise osv.except_osv("Access TOken Expired!", " Please Regenerate Access Token !")

                            else:
                                messages = json.loads((response.content.decode('utf-8')))['value']
                                for message in messages:
                                    if 'from' not in message.keys() or self.env['mail.mail'].search(
                                            [('office_id', '=', message['id'])]) or self.env['mail.message'].search(
                                        [('office_id', '=', message['id'])]):
                                        continue

                                    if 'address' not in message.get('from').get('emailAddress') or message[
                                        'bodyPreview'] == "":
                                        continue

                                    attachment_ids = self.getAttachment(message)

                                    from_partner = self.env['res.partner'].search(
                                        [('email', "=", message['from']['emailAddress']['address'])])
                                    if not from_partner:
                                        continue
                                    from_partner = from_partner[0] if from_partner else from_partner
                                    # if from_partner:
                                    #     from_partner = from_partner[0]
                                    recipient_partners = []
                                    channel_ids = []
                                    for recipient in message['toRecipients']:
                                        if recipient['emailAddress'][
                                            'address'].lower() == res_user.office365_email.lower() or \
                                                recipient['emailAddress'][
                                                    'address'].lower() == res_user.office365_id_address.lower():
                                            to_user = self.env['res.users'].search(
                                                [('id', "=", self._uid)])
                                        else:
                                            to = recipient['emailAddress']['address']
                                            to_user = self.env['res.users'].search(
                                                [('office365_id_address', "=", to)])
                                            to_user = to_user[0] if to_user else to_user
                                        if to_user:
                                            to_partner = to_user.partner_id
                                            recipient_partners.append(to_partner.id)
                                    date = datetime.strptime(message['sentDateTime'], "%Y-%m-%dT%H:%M:%SZ")
                                    self.env['mail.message'].create({
                                        'subject': message['subject'],
                                        'date': date,
                                        'body': message['bodyPreview'],
                                        'email_from': message['from']['emailAddress']['address'],
                                        'partner_ids': [[6, 0, recipient_partners]],
                                        'attachment_ids': [[6, 0, attachment_ids]],
                                        'office_id': message['id'],
                                        'author_id': from_partner.id,
                                        'model': 'res.partner',
                                        'res_id': from_partner.id
                                    })
                                    new_email.append(message['id'])
                                    self.env.cr.commit()
                            if '@odata.nextLink' in json.loads((response.content.decode('utf-8'))):
                                print('6')
                                u = json.loads((response.content.decode('utf-8')))['@odata.nextLink']
                                u_one = url.split('?')[0]
                                u_two = u.split('?')[1].split('10&%24')[1]
                                url = u_one + '?' + u_two

                                # url = response['@odata.nextLink']
                                print('7')

                    response = requests.get(url,
                                            headers={
                                                'Host': 'outlook.office.com',
                                                'Authorization': 'Bearer {0}'.format(res_user.token),
                                                'Accept': 'application/json',
                                                'X-Target-URL': 'http://outlook.office.com',
                                                'connection': 'keep-Alive'
                                            })

                    if 'value' not in json.loads((response.content.decode('utf-8'))).keys():
                        raise osv.except_osv("Access TOken Expired!", " Please Regenerate Access Token !")

                    else:
                        messages = json.loads((response.content.decode('utf-8')))['value']
                        for message in messages:
                            if 'from' not in message.keys() or self.env['mail.mail'].search(
                                    [('office_id', '=', message['id'])]) or self.env['mail.message'].search(
                                [('office_id', '=', message['id'])]):
                                continue

                            if 'address' not in message.get('from').get('emailAddress') or message['bodyPreview'] == "":
                                continue

                            attachment_ids = self.getAttachment(message)

                            from_partner = self.env['res.partner'].search(
                                [('email', "=", message['from']['emailAddress']['address'])])
                            if not from_partner:
                                continue
                            from_partner = from_partner[0] if from_partner else from_partner
                            # if from_partner:
                            #     from_partner = from_partner[0]
                            recipient_partners = []
                            channel_ids = []
                            for recipient in message['toRecipients']:
                                if recipient['emailAddress'][
                                    'address'].lower() == res_user.office365_email.lower() or \
                                        recipient['emailAddress'][
                                            'address'].lower() == res_user.office365_id_address.lower():
                                    to_user = self.env['res.users'].search(
                                        [('id', "=", self._uid)])
                                else:
                                    to = recipient['emailAddress']['address']
                                    to_user = self.env['res.users'].search(
                                        [('office365_id_address', "=", to)])
                                    to_user = to_user[0] if to_user else to_user
                                if to_user:
                                    to_partner = to_user.partner_id
                                    recipient_partners.append(to_partner.id)
                            date = datetime.strptime(message['sentDateTime'], "%Y-%m-%dT%H:%M:%SZ")
                            self.env['mail.message'].create({
                                'subject': message['subject'],
                                'date': date,
                                'body': message['bodyPreview'],
                                'email_from': message['from']['emailAddress']['address'],
                                'partner_ids': [[6, 0, recipient_partners]],
                                'attachment_ids': [[6, 0, attachment_ids]],
                                'office_id': message['id'],
                                'author_id': from_partner.id,
                                'model': 'res.partner',
                                'res_id': from_partner.id
                            })
                            new_email.append(message['id'])
                            self.env.cr.commit()
                    if '@odata.nextLink' in json.loads((response.content.decode('utf-8'))):
                        print('6')
                        u = json.loads((response.content.decode('utf-8')))['@odata.nextLink']
                        u_one = url.split('?')[0]
                        u_two = u.split('?')[1].split('10&%24')[1]
                        url = u_one + '?' + u_two

                        # url = response['@odata.nextLink']
                        print('7')
            except Exception as e:
                # res_user.send_mail_flag = True
                status = 'Error'
                _logger.error(e)
                raise ValidationError(_(str(e)))

    def sync_customer_sent_mail(self):
        context = self._context
        current_uid = context.get('uid')
        res_user = self.env['res.users'].browse(current_uid)
        if res_user.token:
            try:
                if res_user.expires_in:
                    expires_in = datetime.fromtimestamp(int(res_user.expires_in) / 1e3)
                    expires_in = expires_in + timedelta(seconds=3600)
                    nowDateTime = datetime.now()
                    if nowDateTime > expires_in:
                        self.generate_refresh_token()

                response = requests.get(
                    'https://graph.microsoft.com/v1.0/me/mailFolders',
                    headers={
                        'Host': 'outlook.office.com',
                        'Authorization': 'Bearer {0}'.format(res_user.token),
                        'Accept': 'application/json',
                        'X-Target-URL': 'http://outlook.office.com',
                        'connection': 'keep-Alive'
                    }).content
                if 'value' not in json.loads((response.decode('utf-8'))).keys():
                    raise osv.except_osv("Access Token Expired!", " Please Regenerate Access Token !")
                else:
                    folders = json.loads((response.decode('utf-8')))['value']
                    sentbox_folder_id = [folder['id'] for folder in folders if folder['displayName'] == 'Sent Items']
                    if sentbox_folder_id:
                        sentbox_id = sentbox_folder_id[0]
                        response = requests.get(
                            'https://graph.microsoft.com/v1.0/me/mailFolders/' + sentbox_id + '/messages?$top=100000&$count=true',
                            headers={
                                'Host': 'outlook.office.com',
                                'Authorization': 'Bearer {0}'.format(res_user.token),
                                'Accept': 'application/json',
                                'X-Target-URL': 'http://outlook.office.com',
                                'connection': 'keep-Alive'
                            }).content
                        if 'value' not in json.loads((response.decode('utf-8'))).keys():

                            raise osv.except_osv("Access Token Expired!", " Please Regenerate Access Token !")
                        else:
                            messages = json.loads((response.decode('utf-8')))['value']
                            for message in messages:
                                # print(message['bodyPreview'])

                                if 'from' not in message.keys() or self.env['mail.mail'].search(
                                        [('office_id', '=', message['conversationId'])]) or self.env[
                                    'mail.message'].search(
                                    [('office_id', '=', message['conversationId'])]):
                                    continue

                                if message['bodyPreview'] == "":
                                    continue

                                attachment_ids = self.getAttachment(message)
                                if message['from']['emailAddress'][
                                    'address'].lower() == res_user.office365_email.lower() or \
                                        message['from']['emailAddress'][
                                            'address'].lower() == res_user.office365_id_address.lower():
                                    email_from = res_user.email
                                else:
                                    email_from = message['from']['emailAddress']['address']

                                from_user = self.env['res.users'].search(
                                    [('id', "=", self._uid)])
                                if from_user:
                                    from_partner = from_user.partner_id
                                else:
                                    continue

                                channel_ids = []
                                for recipient in message['toRecipients']:

                                    to_partner = self.env['res.partner'].search(
                                        [('email', "=", recipient['emailAddress']['address'])])
                                    to_partner = to_partner[0] if to_partner else to_partner

                                    if not to_partner:
                                        continue
                                    date = datetime.strptime(message['sentDateTime'], "%Y-%m-%dT%H:%M:%SZ")
                                    self.env['mail.message'].create({
                                        'subject': message['subject'],
                                        'date': date,
                                        'body': message['bodyPreview'],
                                        'email_from': email_from,
                                        'partner_ids': [[6, 0, [to_partner.id]]],
                                        'attachment_ids': [[6, 0, attachment_ids]],
                                        'office_id': message['conversationId'],
                                        'author_id': from_partner.id,
                                        'model': 'res.partner',
                                        'res_id': to_partner.id
                                    })
                                    self.env.cr.commit()

            except Exception as e:
                _logger.error(e)
                raise ValidationError(_(str(e)))

    def getAttachment(self, message):
        context = self._context
        current_uid = context.get('uid')
        res_user = self.env['res.users'].browse(current_uid)
        if res_user.expires_in:
            expires_in = datetime.fromtimestamp(int(res_user.expires_in) / 1e3)
            expires_in = expires_in + timedelta(seconds=3600)
            nowDateTime = datetime.now()
            if nowDateTime > expires_in:
                self.generate_refresh_token()

        response = requests.get(
            'https://graph.microsoft.com/v1.0/me/messages/' + message['id'] + '/attachments/',
            headers={
                'Host': 'outlook.office.com',
                'Authorization': 'Bearer {0}'.format(res_user.token),
                'Accept': 'application/json',
                'X-Target-URL': 'http://outlook.office.com',
                'connection': 'keep-Alive'
            }).content
        attachments = json.loads((response.decode('utf-8')))['value']
        attachment_ids = []
        for attachment in attachments:
            if 'contentBytes' not in attachment or 'name' not in attachment:
                continue
            odoo_attachment = self.env['ir.attachment'].create({
                'datas': attachment['contentBytes'],
                'name': attachment["name"],
                'store_fname': attachment["name"]})
            self.env.cr.commit()
            attachment_ids.append(odoo_attachment.id)
        return attachment_ids

    def export_contacts(self):

        context = self._context
        current_uid = context.get('uid')
        res_user = self.env['res.users'].browse(current_uid)
        new_contact = []
        update_contact = []
        status = None
        if res_user.token:
            try:
                if res_user.token:
                    if res_user.expires_in:
                        expires_in = datetime.fromtimestamp(int(res_user.expires_in) / 1e3)
                        expires_in = expires_in + timedelta(seconds=3600)
                        nowDateTime = datetime.now()
                        if nowDateTime > expires_in:
                            self.generate_refresh_token()

                    # odoo_contacts = self.env['res.partner'].search(
                    #     ['|',('company_id', '=', res_user.company_id.id), ('company_id', '=', None)])

                    odoo_contacts = self.env['res.partner'].search(
                        [('office_contact_id', '=', None)])

                    office_contact = []
                    count = 0
                    if odoo_contacts:
                        url_count = 'https://graph.microsoft.com/beta/me/contacts?$count = true'

                        headers = {

                            'Host': 'outlook.office365.com',
                            'Authorization': 'Bearer {0}'.format(res_user.token),
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-Target-URL': 'http://outlook.office.com',
                            'connection': 'keep-Alive'

                        }

                        response_count = requests.get(
                            url_count, headers=headers
                        ).content

                        response_count = json.loads(response_count.decode('utf-8'))
                        if '@odata.count' in response_count and response_count['@odata.count'] != -1:
                            count = response_count['@odata.count']

                        url = 'https://graph.microsoft.com/v1.0/me/contacts?$top=' + str(count)

                        response = requests.get(
                            url, headers=headers
                        ).content
                        response = json.loads(response.decode('utf-8'))
                        if not 'value' in response:
                            raise osv.except_osv("Access Token Expired!", " Please Regenerate Access Token !")

                        if 'value' in response:
                            contacts_emails = [response['value'][i]['emailAddresses'] for i in
                                               range(len(response['value']))]
                            for cont in contacts_emails:
                                if cont:
                                    office_contact.append(cont[0]['address'])

                        for contact in odoo_contacts:
                            company = None

                            if contact.company_name:
                                company = contact.company_name
                            elif contact.parent_id.name:
                                company = contact.parent_id.name

                            data = {
                                "givenName": contact.name if contact.name else None,
                                'companyName': company,
                                'mobilePhone': contact.mobile if contact.mobile else None,
                                'jobTitle': contact.function if contact.function else None,
                                "businessPhones": [
                                    contact.phone if contact.phone else None
                                ]
                            }
                            if contact.email:
                                data["emailAddresses"] = [
                                    {
                                        "address": contact.email,
                                    }
                                ]
                            data["homeAddress"]= {
                                "street": contact.street if contact.street else (contact.street2 if contact.street2 else None),
                                "city": contact.city if contact.city else None,
                                "state": contact.state_id.name if contact.state_id else None,
                                "countryOrRegion": contact.country_id.name if contact.country_id else None,
                                "postalCode": contact.zip if contact.zip else None
                            }
                            if not contact.email and not contact.mobile and not contact.phone:
                                continue
                            if contact.office_contact_id or contact.email in office_contact:
                                if contact.create_date < contact.write_date and contact.is_update:
                                    if contact.office_contact_id:
                                        update_response = requests.patch(
                                            'https://graph.microsoft.com/v1.0/me/contacts/'+str(contact.office_contact_id), data=json.dumps(data), headers=headers
                                        )
                                        if update_response.status_code != 200:
                                            pass

                                        else:
                                            response = json.loads(update_response.content)
                                            contact.write({'office_contact_id': response['id']})
                                            update_contact.append(response['id'])
                                        contact.is_update=False
                                else:
                                    continue

                            else:

                                post_response = requests.post(
                                    'https://graph.microsoft.com/v1.0/me/contacts', data=json.dumps(data), headers=headers
                                ).content

                                if 'id' not in json.loads(post_response.decode('utf-8')).keys():
                                    raise osv.except_osv(_("Error!"), (_(post_response["error"])))
                                else:
                                    response = json.loads(post_response.decode('utf-8'))
                                    contact.write({'office_contact_id': response['id']})
                                    contact.is_update = False
                                    new_contact.append(response['id'])

                else:
                    raise UserWarning('Token is missing. Please Generate Token ')

            except Exception as e:
                _logger.error(e)
                status = "Error"
                raise ValidationError(_(str(e)))

class CustomContacts(models.Model):

    _inherit = 'res.partner'

    office_contact_id = fields.Char('Office365 Id')
    is_update = fields.Boolean(string="Is update",default=True)
    modified_date = fields.Datetime('Modified Date')
    is_create = fields.Boolean(string="Is create", default=True  )

    # @api.model
    # def write(self,values):
    #     if 'is_update' in values:
    #         return super(CustomContacts, self).write(values)
    #     else:
    #         if 'office_contact_id' not in values:
    #             values['is_update']=True
    #         return super(CustomContacts, self).write(values)
    #
    # @api.model
    # def create(self, values):
    #     # Add code here
    #
    #     if 'is_create' in values:
    #         return super(CustomContacts, self).create(values)
    #     else:
    #         # return super(CustomContacts, self).create(values)
    #         if 'office_contact_id' not in values:
    #             context = self._context
    #             current_uid = context.get('uid')
    #             res_user = self.env['res.users'].browse(current_uid)
    #             new_contact = []
    #             update_contact = []
    #             status = None
    #             if res_user.token:
    #                 try:
    #                     if res_user.token:
    #                         if res_user.expires_in:
    #                             expires_in = datetime.fromtimestamp(int(res_user.expires_in) / 1e3)
    #                             expires_in = expires_in + timedelta(seconds=3600)
    #                             nowDateTime = datetime.now()
    #                             if nowDateTime > expires_in:
    #                                 Office365Integration.generate_refresh_token(self)
    #
    #                         odoo_contacts = self.env['res.partner'].search(
    #                             ['|', ('company_id', '=', res_user.company_id.id), ('company_id', '=', None)])
    #
    #                         office_contact = []
    #                         count = 0
    #                         if odoo_contacts:
    #                             url_count = 'https://graph.microsoft.com/beta/me/contacts?$count = true'
    #
    #                             headers = {
    #
    #                                 'Host': 'outlook.office365.com',
    #                                 'Authorization': 'Bearer {0}'.format(res_user.token),
    #                                 'Accept': 'application/json',
    #                                 'Content-Type': 'application/json',
    #                                 'X-Target-URL': 'http://outlook.office.com',
    #                                 'connection': 'keep-Alive'
    #
    #                             }
    #
    #                             response_count = requests.get(
    #                                 url_count, headers=headers
    #                             ).content
    #
    #                             response_count = json.loads(response_count.decode('utf-8'))
    #                             if '@odata.count' in response_count and response_count['@odata.count'] != -1:
    #                                 count = response_count['@odata.count']
    #
    #                             url = 'https://graph.microsoft.com/v1.0/me/contacts?$top=' + str(count)
    #
    #                             response = requests.get(
    #                                 url, headers=headers
    #                             ).content
    #                             response = json.loads(response.decode('utf-8'))
    #                             if not 'value' in response:
    #                                 raise osv.except_osv("Access Token Expired!", " Please Regenerate Access Token !")
    #
    #                             if 'value' in response:
    #                                 contacts_emails = [response['value'][i]['emailAddresses'] for i in
    #                                                    range(len(response['value']))]
    #                                 for cont in contacts_emails:
    #                                     if cont:
    #                                         office_contact.append(cont[0]['address'])
    #
    #                             contact =super(CustomContacts, self).create(values)
    #
    #                             # for contact in odoo_contacts:
    #                             company = None
    #
    #                             if contact.company_name:
    #                                 company = contact.company_name
    #                             elif contact.parent_id.name:
    #                                 company = contact.parent_id.name
    #
    #                             data = {
    #                                 "givenName": contact.name if contact.name else None,
    #                                 'companyName': company,
    #                                 'mobilePhone': contact.mobile if contact.mobile else None,
    #                                 'jobTitle': contact.function if contact.function else None,
    #                                 "businessPhones": [
    #                                     contact.phone if contact.phone else None
    #                                 ]
    #                             }
    #                             if contact.email:
    #                                 data["emailAddresses"] = [
    #                                     {
    #                                         "address": contact.email,
    #                                     }
    #                                 ]
    #                             data["homeAddress"] = {
    #                                 "street": contact.street if contact.street else (
    #                                     contact.street2 if contact.street2 else None),
    #                                 "city": contact.city if contact.city else None,
    #                                 "state": contact.state_id.name if contact.state_id else None,
    #                                 "countryOrRegion": contact.country_id.name if contact.country_id else None,
    #                                 "postalCode": contact.zip if contact.zip else None
    #                             }
    #                             if contact.office_contact_id or contact.email in office_contact:
    #                                 if contact.create_date < contact.write_date and contact.is_update:
    #                                     if contact.office_contact_id:
    #                                         update_response = requests.patch(
    #                                             'https://graph.microsoft.com/v1.0/me/contacts/' + str(
    #                                                 contact.office_contact_id), data=json.dumps(data),
    #                                             headers=headers
    #                                         )
    #                                         if update_response.status_code != 200:
    #                                             pass
    #
    #                                         else:
    #                                             response = json.loads(update_response.content)
    #                                             contact['office_contact_id'] = response['id']
    #                                             self.env.cr.commit()
    #                                             # contact.write({'office_contact_id': response['id']})
    #                                             # update_contact.append(response['id'])
    #                                         # contact.is_update = False
    #                             else:
    #                                 post_response = requests.post(
    #                                     'https://graph.microsoft.com/v1.0/me/contacts', data=json.dumps(data),
    #                                     headers=headers
    #                                 ).content
    #                                 if 'id' not in json.loads(post_response.decode('utf-8')).keys():
    #                                     raise osv.except_osv(_("Error!"), (_(post_response["error"])))
    #                                 else:
    #                                     response = json.loads(post_response.decode('utf-8'))
    #                                     contact['office_contact_id'] = response['id']
    #                                     self.env.cr.commit()
    #                                     return contact
    #                                     # contact.write({'office_contact_id': response['id']})
    #                                     # contact.is_update = False
    #                                     # new_contact.append(response['id'])
    #                     else:
    #                         raise UserWarning('Token is missing. Please Generate Token ')
    #
    #                 except Exception as e:
    #                     _logger.error(e)
    #                     status = "Error"
    #                     raise ValidationError(_(str(e)))




class CustomMessageInbox(models.Model):
    _inherit = 'mail.message'
    office_id = fields.Char('Office Id')

class CustomMessage(models.Model):
    _inherit = 'mail.mail'
    office_id = fields.Char('Office Id')

    @api.model
    def create(self, values):
        o365_id = None
        conv_id = None
        context = self._context
        user = self.env['res.users'].browse(self.env.uid)
        # if user.send_mail_flag:
        if user.token:
            if user.expires_in:
                expires_in = datetime.fromtimestamp(int(user.expires_in) / 1e3)
                expires_in = expires_in + timedelta(seconds=3600)
                nowDateTime = datetime.now()
                if nowDateTime > expires_in:
                    self.generate_refresh_token()
            if 'mail_message_id' in values:
                email_obj = self.env['mail.message'].search([('id', '=', values['mail_message_id'])])
                partner_id = values['recipient_ids'][0][1]
                partner_obj = self.env['res.partner'].search([('id', '=', partner_id)])

                new_data = {
                            "subject": values['subject'] if values['subject'] else email_obj.body,
                            # "importance": "high",
                            "body": {
                                "contentType": "HTML",
                                "content": email_obj.body
                            },
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        "address": partner_obj.email
                                    }
                                }
                            ]
                        }

                response = requests.post(
                    'https://graph.microsoft.com/v1.0/me/messages', data=json.dumps(new_data),
                                        headers={
                                            'Host': 'outlook.office.com',
                                            'Authorization': 'Bearer {0}'.format(user.token),
                                            'Accept': 'application/json',
                                            'Content-Type': 'application/json',
                                            'X-Target-URL': 'http://outlook.office.com',
                                            'connection': 'keep-Alive'
                                        })
                if 'conversationId' in json.loads((response.content.decode('utf-8'))).keys():
                    conv_id = json.loads((response.content.decode('utf-8')))['conversationId']

                if 'id' in json.loads((response.content.decode('utf-8'))).keys():

                    o365_id = json.loads((response.content.decode('utf-8')))['id']
                    if email_obj.attachment_ids:
                        for attachment in self.getAttachments(email_obj.attachment_ids):
                            attachment_response = requests.post(
                                'https://graph.microsoft.com/beta/me/messages/' + o365_id + '/attachments',
                                data=json.dumps(attachment),
                                headers={
                                    'Host': 'outlook.office.com',
                                    'Authorization': 'Bearer {0}'.format(user.token),
                                    'Accept': 'application/json',
                                    'Content-Type': 'application/json',
                                    'X-Target-URL': 'http://outlook.office.com',
                                    'connection': 'keep-Alive'
                                })
                    send_response = requests.post(
                        'https://graph.microsoft.com/v1.0/me/messages/' + o365_id + '/send',
                        headers={
                            'Host': 'outlook.office.com',
                            'Authorization': 'Bearer {0}'.format(user.token),
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-Target-URL': 'http://outlook.office.com',
                            'connection': 'keep-Alive',
                            'Content-Length': '0'
                        })

                    message = super(CustomMessage, self).create(values)
                    message.email_from = None

                    if conv_id:
                        message.office_id = conv_id

                    return message
                else:
                    pass

            else:

                return super(CustomMessage, self).create(values)

        else:
            return super(CustomMessage, self).create(values)
    def getAttachments(self, attachment_ids):
        attachment_list = []
        if attachment_ids:
            # attachments = self.env['ir.attachment'].browse([id[0] for id in attachment_ids])
            attachments = self.env['ir.attachment'].search([('id', 'in', [i.id for i in attachment_ids])])
            for attachment in attachments:
                attachment_list.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.name,
                    "contentBytes": attachment.datas.decode("utf-8")
                })
        return attachment_list

    def generate_refresh_token(self):
        context = self._context
        current_uid = context.get('uid')
        # res_user = self.env['res.users'].browse(current_uid)
        user = self.env['res.users'].browse(current_uid)
        if user.refresh_token:
            # settings = self.env['res.users'].search([])
            # settings = settings[0] if settings else settings

            if not user.client_id or not user.redirect_url or not user.secret_id:
                raise osv.except_osv(_("Error!"), (_("Please ask admin to add Office365 settings!")))
            header = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }


            response = requests.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                data='grant_type=refresh_token&refresh_token=' + user.refresh_token + '&redirect_uri=' + user.redirect_url + '&client_id=' + user.client_id + '&client_secret=' + user.secret_id
                , headers=header).content

            response = json.loads((str(response)[2:])[:-1])
            if 'access_token' not in response:
                response["error_description"] = response["error_description"].replace("\\r\\n", " ")
                raise osv.except_osv(_("Error!"), (_(response["error"] + " " + response["error_description"])))
            else:
                user.token = response['access_token']
                user.refresh_token = response['refresh_token']
                user.expires_in = int(round(time.time() * 1000))
                self.env.cr.commit()
