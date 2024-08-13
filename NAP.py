import logging.config
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
from unidecode import unidecode
from dotenv import load_dotenv
import os
import requests
from io import BytesIO
import json
import base64
import asyncio
from azure.identity.aio import ClientSecretCredential
from dateutil.relativedelta import relativedelta
import logging
import colorlog


class MsGraph:
    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
            
    async def generate_msgraph_token(self):
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )
        scopes = ["https://graph.microsoft.com/.default"]
        token = await credential.get_token(*scopes)
        return token.token
    
    async def generate_msgraph_headers(self):
        token = await self.generate_msgraph_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        return headers
        
async def msgraph_sdk_menager(tenant_id, client_id, client_secret):
    sdk = MsGraph(tenant_id, client_id, client_secret)
    headers = await sdk.generate_msgraph_headers()
    return headers

async def get_user(tenant_id, client_id, client_secret, user_id):
    headers = await msgraph_sdk_menager(tenant_id, client_id, client_secret)
    endpoint = f'https://graph.microsoft.com/v1.0/users/{user_id}'
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        user = response.json()
        print(user)
    else:
        print(f'Error: {response.status_code}, {response.text}')

async def msgraph_main(tenant_id, client_id, client_secret, user_id):
    await get_user(tenant_id, client_id, client_secret, user_id)
    
class SharepointData():

    def __init__(self, headers):
        self.headers = headers
        
    def get_sharepoint_newbies_credentials(self, site_id, newbies_credentials_list_id):
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{newbies_credentials_list_id}/items?$expand=fields'
        response = requests.get(url, headers=self.headers)
        if response.status_code != 200:
            raise Exception(response.json())
        data = response.json()
        # iterate over json 
        employee_data = []
        for item in data['value']:
            fields = item['fields']
            employee_id = fields['Title']
            entra_id = fields['AzADObjectId']
            onepassword_link = fields['PasswordShareLink']
            employee_data.append([employee_id,entra_id,onepassword_link])
        return employee_data
          
class SuluData():

    def __init__(self, application_id, headers):
        self.application_id = application_id
        self.headers = headers
        
    def get_sulu_data(self, employee_id):
        url = f"https://graph.microsoft.com/v1.0/users?$filter=employeeId eq '{employee_id}'"
        response = requests.get(url, headers=self.headers)
        if response.status_code != 200:
            raise Exception(response.json())
        result = response.json()
        for value in result['value']:
            self.microsoft_id = value['id']
        extensions_url = f'https://graph.microsoft.com/v1.0/applications/{self.application_id}/extensionProperties'
        extensions_response = requests.get(extensions_url, headers=self.headers)
        extension_properties = extensions_response.json().get('value', [])
        user_properties = [
            'id', 'displayName', 'userPrincipalName', 'userType', 'createdDateTime', 'accountEnabled','name','menager',
            'onPremisesDistinguishedName', 'onPremisesSyncEnabled',
            'licenseAssignmentStates',
            'signInActivity',
            'employeeId', 'jobTitle', 'companyName', 'mail',
            'sponsors'
        ]
        for prop in extension_properties:
            if 'User' in prop.get('targetObjects', []):
                user_properties.append(prop.get('name'))
        user_properties_str = ','.join(user_properties)
        user_id = self.microsoft_id
        user_url = f'https://graph.microsoft.com/v1.0/users/{user_id}'
        user_response = requests.get(user_url, headers=self.headers, params={'$select': user_properties_str})
        self.user = user_response.json()
        return self.user

class Newcomers():
    def __init__(self):
        self.indexes = []
        self.office = pd.DataFrame()  
        self.mails = pd.DataFrame()   
        self.self_pickup = pd.DataFrame()  
        self.valid_addresses = []
        self.key = os.getenv("ADDRESS_VALIDATION_API_KEY_LINGARO")
        
    def validate_address(self,api_key,address_to_validate):
        logger.info("Method - validate_address init.")
        url = "https://addressvalidation.googleapis.com/v1:validateAddress"
        params = {
            "key" : api_key
        }
        payload = {
            "address": {
                "address_lines" : address_to_validate
            }
        }
        headers = {
                "Content-Type" : "application/json"
        }
        response = requests.post(url=url, headers=headers, params=params, data=json.dumps(payload))

        if response.status_code != 200:
            raise Exception(response.json())
        result = response.json()
        formatted_address = result['result']['address']['formattedAddress']  # address after validation
        return formatted_address
    
    def get_excel_file_from_sharepoint(self, drive_id, item_id, headers):
        endpoint = 'https://graph.microsoft.com/v1.0'
        response = requests.get(
            url=endpoint + f"/drives/{drive_id}/items/{item_id}",
            headers=headers
        )
        if response.status_code != 200:
            raise Exception(f"Failed to fetch SharePoint file: {response.json()}")
        result = response.json()
        download_url = result['@microsoft.graph.downloadUrl']
        download_response = requests.get(url=download_url)
        download_response.raise_for_status()
        self.file_content = BytesIO(download_response.content)
        excel_file = pd.ExcelFile(self.file_content)
        sheet_names = excel_file.sheet_names
        logger.debug(f"Method get_excel_file_from_sharepoint \n {sheet_names}")
        return sheet_names

    def create_dataframe(self, sheet):
        logger.info(f"Method create_dataframe init.")
        df = pd.read_excel(self.file_content, sheet_name=sheet)
        logger.debug(f"Raw DataFeame from create_datafeame method: {df.to_string()}")
        return df
    def filter_dataframe(self,raw_df):
        logger.info()
    def clean_newcomers_excel_data(self, raw_df):
        logger.info(f"Method clean_newcomers_excel_data init.")
        data = ['employeeID', 'name', 'address', 'phone', 'start date', 'e-mail before start', 'laptop','telefon sluzbowy', 'umowa', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']
        filt = (~raw_df['address'].str.contains("Mexico|MEXICO|México", na=False))
        raw_df = raw_df.loc[filt, data]
        raw_df = raw_df.dropna(subset = ['name'])
        raw_df['name'] = raw_df['name'].apply(unidecode).str.strip().str.lower()
        raw_df = raw_df.dropna(subset=['address'])
        raw_df = raw_df.dropna(subset=['employeeID'])
        raw_df.drop(raw_df[raw_df['umowa'] != "podpisana"].index, inplace = True)
        raw_df['laptop'] = raw_df['laptop'].replace(np.nan, "standard win" ,regex = True)
        raw_df['employeeID'] = raw_df['employeeID'].astype(int)
        raw_df['telefon sluzbowy'] = raw_df['telefon sluzbowy'].replace(np.nan, " " ,regex = True)
        raw_df['phone'] = raw_df['phone'].astype(str).str.replace(" ", "")
        raw_df['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'] = raw_df['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'].replace(np.nan, " ", regex = True)
        logger.debug(f"Cleaned DataFrame: \n {raw_df.to_string()}")
        return raw_df

    def calculate_days_to_start(self, clean_df):
        logger.info(f"Method calculate_days_to_start init.")
        current_date = datetime.today()
        
        for index, row in clean_df.iterrows():
            start_date = row['start date'] 
            logger.info(f"{row['name']} | {start_date} | {type(start_date)}")
            if not isinstance(start_date, datetime):
                start_date = datetime.strptime(start_date, '%d.%m.%Y')
                logger.warning(f"Wrong start date format was found in: \n {row}")
            if current_date <= start_date <= current_date + timedelta(days=25) and start_date.weekday() in [0, 1, 5, 6]:  
                self.indexes.append(int(index))  
            elif start_date - timedelta(days=23) <= current_date <= start_date and start_date.weekday() in [2, 3, 4]: 
                self.indexes.append(int(index))
                
        couple_days_away = clean_df.loc[self.indexes]
        df = couple_days_away['address'].values.tolist()
        
        for address in df:
            validate = self.validate_address(self.key,address)
            self.valid_addresses.append(validate)
            
        couple_days_away['address'] = self.valid_addresses
        self.equpiment_data = couple_days_away[['employeeID','name', 'start date', 'laptop', 'telefon sluzbowy', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
        if not couple_days_away['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'].str.strip().str.lower().eq("osobiście odbiór".strip().lower()).any():
            self.office = couple_days_away[['employeeID','name', 'address', 'phone']]
            self.mails = couple_days_away[['e-mail before start']]
        else:
            self.self_pickup = couple_days_away[['employeeID','name', 'address','phone','Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
        return couple_days_away

    def get_equipment_data(self):
        return self.equpiment_data
    
    def get_employee_personal_mail(self):
        return self.mails

    def get_office_pickup_list(self):
        if self.self_pickup.empty:
            return 0
        else:
            return self.self_pickup
        
    def get_courier_shippment(self):
        return self.office

class MailSender():
    def draft_atttachment(self,file_path):
        if not os.path.exists(file_path):
            print('File is not found')
            return None
        with open(file_path, 'rb') as upload:
            media_content = base64.b64encode(upload.read())
        data_body = {
            '@odata.type': '#microsoft.graph.fileAttachment',
            'contentBytes': media_content.decode('utf-8'),
            'name': os.path.basename(file_path)
        }
        return data_body

    def mail_body(self,address, subject, content, attachment):
        requests_body = {
            'message': {
                'toRecipients': [
                    {
                        'emailAddress': {
                            'address': address#
                        }
                    }
                ],
                'bccRecipients': [
                    {
                        'emailAddress': {
                            'address': "maciej.cichocki@lingarogroup.com" 
                        }
                    }
                ],
                'subject': subject,
                'importance': 'normal',
                'body': {
                    'contentType': 'HTML',
                    'content': f"<b>{content}</b>"
                },
                'attachments': []
            }
        }
        if attachment and os.path.exists(attachment):
            attachment = self.draft_atttachment(file_path=attachment)
            if attachment:
                requests_body['message']['attachments'].append(attachment)
        return requests_body
        
    def send_mail(self,user_id,address, subject, content, attachment, headers):
        GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'
        endpoint = GRAPH_ENDPOINT + f"/users/{user_id}/sendMail"
        mail = self.mail_body(address, subject, content, attachment)
        response = requests.post(endpoint, headers=headers, json=mail)
        if response.status_code != 202:
            raise Exception(response.json())
        print(f"Mail was send to - {address}")

    def send_welcome_mail_to_newcomer(self, employee_full_name, employee_start_date, onepassword_link, user_personal_mail):
        employee_start_date_with_time = datetime.combine(employee_start_date, time(hour=8, minute=0))
        employee_STARTDATE = employee_start_date_with_time.strftime("%d %B %Y %H:%M")
        employee_ONEPASSWORD_ACTIVE_DATE = datetime.combine(
            employee_start_date_with_time - timedelta(days=2),
            time(hour=4, minute=0)
        )
        employee_CREDENTIALS_ENABLED = employee_ONEPASSWORD_ACTIVE_DATE.strftime("%d %B %Y %H:%M")
        employee_name = employee_full_name.split()
        payload = json.dumps({
        "Personalizations": [
            {
            "To": [
                {
                "Email": user_personal_mail
                }
            ],
            "Bcc": [
                {
                "Email": "maciej.cichocki@lingarogroup.com"
                }
            ],
            "dynamic_template_data": {
                "Employee": {
                "FirstDayOfWork": employee_STARTDATE,
                "OfficeCountryCode": "PL",
                "OnePasswordUrl": onepassword_link,
                "FirstName": employee_name[0],
                "AccountEnabledFrom": employee_CREDENTIALS_ENABLED
                },
                "Assets": {
                "Delivery": {
                    "UseDelivery": False
                }
                }
            }
            }
        ],
        "From": {
            "Email": "no-reply@lingaro.io",
            "Name": "Lingaro"
        },
        "template_id": "d-3f208ee9afdc4a79a80337673c228a56"
        })
        headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + os.getenv("SEND_GRID_CREDENTIALS")
        }
        url = url = "https://api.sendgrid.com/v3/mail/send"
        response = requests.post(url, headers=headers, data=payload)
        
        if response.status_code != 202:
            raise Exception(response.text)
        response.text
        print("Mail was Send to - {} ".format(employee_full_name))
    
class NewcomersManager():
    def __init__(self, drive_id, item_id, headers, sheet):
        self.newcomers = Newcomers()
        self.excel_sheets_list = self.newcomers.get_excel_file_from_sharepoint(drive_id, item_id, headers) 
        self.raw_dataframe = self.newcomers.create_dataframe(sheet)
        self.clean_dataframe = self.newcomers.clean_newcomers_excel_data(raw_dataframe=self.raw_dataframe)
        self.processed_dataframe = self.newcomers.calculate_days_to_start(clean_df=self.clean_dataframe)
    
    def get_excel_data(self):
        return self.processed_dataframe
    
    def get_shipment_data(self):
        return self.newcomers.get_courier_shippment()

    def get_office_pickup_data(self):
        return self.newcomers.get_office_pickup_list()

    def get_equipment_data(self):
        return self.newcomers.get_equipment_data()
    
class Dhl():
    """Connect via api and create label for shipping, get track number"""
class Jira():
    """Figure out how to connect via api, find new person tickets and close them"""
def setup_logger():
    logging_config = {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {
            "detailed": {
                "()": colorlog.ColoredFormatter,
                "format": "\n[%(log_color)s%(levelname)s|%(module)s|L%(lineno)d] %(asctime)s: %(log_color)s%(message)s",
                "datefmt": "%Y-%m-%dT%H:%M:%S%z",
                "log_colors": {
                    "DEBUG": "cyan",
                    "INFO": "green",
                    "WARNING": "yellow",
                    "ERROR": "red",
                    "CRITICAL": "red,bg_white",
                }
            }
        },
        "handlers": {
            "console": {
                "class": "logging.StreamHandler",
                "level": "DEBUG",
                "formatter": "detailed"
            }
        },
        "root": {
            "level": "DEBUG",
            "handlers": ["console"]
        }
    }
    my_logger = logging.config.dictConfig(logging_config)
    logger = logging.getLogger(my_logger)  # create a logger with a custom name
    return logger     
      
def process_string(s):
    s = unidecode(s)
    s = s.strip()
    s = s.lower()
    return s

def get_mail_sender_instance():
    return MailSender()    

def get_sharepoint_data(headers,site_id, newbies_credentials_list_id):
    sharepoint_data = SharepointData(headers)
    data = sharepoint_data.get_sharepoint_newbies_credentials(site_id, newbies_credentials_list_id)
    return data

def get_sulu_data(application_id, headers,employee_id):
    sulu_data = SuluData(application_id, headers)
    data = sulu_data.get_sulu_data(employee_id)
    return data

def get_excel_sheet(drive_id, item_id, headers):
        logger.info("func - get_excel_sheet init.")
        newcomers = Newcomers()
        sheet_list = newcomers.get_excel_file_from_sharepoint(drive_id, item_id, headers)
        month_list = []
        current_date = datetime.now()
        current_date = datetime.strftime(current_date, "%B %Y").lower().strip()
        sheet_uniform = []
        for sheet in sheet_list:
            sheet_uniform.append(sheet.lower().strip())
        next_month = datetime.now() + relativedelta(months=1)
        next_month = next_month.replace(day=1)
        next_month = datetime.strftime(next_month, "%B %Y").lower().strip()
        time_period = datetime.now() + timedelta(days=6)
        time_period = datetime.strftime(time_period, "%B %Y").lower().strip()
        if time_period >= next_month: 
            month_list.append(current_date)
            month_list.append(next_month)
        else:
            month_list.append(current_date)
            
        return month_list, sheet_uniform
    
def get_newcomers_data(drive_id, item_id, headers, sheet): 
    menager = NewcomersManager(drive_id, item_id, headers, sheet)
    newcomers_excel_data = menager.get_excel_data()
    shipment_data = menager.get_shipment_data()
    office_pickup = menager.get_office_pickup_data()
    equipment_data = menager.get_equipment_data()
    return newcomers_excel_data, shipment_data, office_pickup , equipment_data

def process_employee_data(sulu_data,newcomers_excel_data, sharepoint_data, mail_sender, headers, user_id):
    employee_data_for_sharepoint_email_tracking_list = set()
    employes_from_ltl = set()
    for index,row in newcomers_excel_data.iterrows():
        excel_employee_id = row["employeeID"]
        excel_employee_name = row["name"]
        excel_employee_start_date = row["start date"]
        try:
            sulu_employee_data = sulu_data.get_sulu_data(excel_employee_id)
            sulu_employee_name = sulu_employee_data['displayName']
            if excel_employee_name == process_string(sulu_employee_name):
                sulu_employee_id = sulu_employee_data["id"]
                excel_employee_start_date = str(row["start date"])
                excel_employee_personal_mail = row["e-mail before start"]
                sharepoint_found = False
                for sharepoint in sharepoint_data:
                    if int(sharepoint[0]) == excel_employee_id:
                        employee_data_for_sharepoint_email_tracking_list.add((
                            sulu_employee_id,
                            sulu_employee_name,
                            excel_employee_id,
                            excel_employee_start_date,
                            excel_employee_personal_mail,
                            sharepoint[2],
                        ))
                        sharepoint_found = True
                        break
                if not sharepoint_found:
                    employee_data_for_sharepoint_email_tracking_list.add((
                        sulu_employee_id,
                        sulu_employee_name,
                        excel_employee_id,
                        excel_employee_start_date,
                        excel_employee_personal_mail,
                        "https://share.1password.com/s#rcAv4wgslR3cUc--7JFCR935dD-veFGcrF7pXpxoRXc"
                    ))
                    employes_from_ltl.add((
                        excel_employee_id,
                        sulu_employee_id,
                        sulu_employee_name,
                        excel_employee_start_date
                    ))
 
        except AttributeError as e:
            print(f"Attribute Error {e}")
    if employes_from_ltl:
        df = pd.DataFrame(employes_from_ltl, columns=['Employee id','Microsoft id', 'Name', 'Start date'])
        content = f"Employee password must be reset to - L1n99aROrba22 <br> {df.to_html(index=False)}"
        mail_sender.send_mail(user_id,"dominik.boras@lingarogroup.com", "Long Term Leavers", content, "2137", headers)
    print(employes_from_ltl)
    return employee_data_for_sharepoint_email_tracking_list # employee_data

def add_sharepoint_email_tracking_record(site_id, email_tracking_list_id, headers, employee_data, mail_sender, shippment_data, office_pick_up, equipment_data, user_if):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{email_tracking_list_id}/items"
    for data in employee_data:
        entra_id = str(data[0])
        employee_name = str(data[1])
        employee_id = str(data[2])
        employee_start_date = str(data[3])  # format = '2024-07-22 00:00:00'
        employee_personal_mail = str(data[4])
        employee_one_password = str(data[5])
        start_date = datetime.strptime(employee_start_date, "%Y-%m-%d %H:%M:%S")
        end_date = (start_date + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S")
        payload = {
                "fields": {
                    "Title": f"IT Welcome - {employee_name}",
                    "EntraId": entra_id,
                    "EmployeeId": employee_id,
                    "SendGridTemplateId": "d-3f208ee9afdc4a79a80337673c228a56",
                    "Category": "Employee Lifecycle",
                    "SubCategory": "Pre-Onboarding",               
                    "ExpirationDate": end_date,
            }
        }
        response = requests.post(url, headers=headers, json=payload)
        if response.status_code != 201:
            raise Exception(response.json())
        mail_sender.send_welcome_mail_to_newcomer(employee_name, start_date, employee_one_password, employee_personal_mail)
     
    shippment_data = shippment_data[['name','address','phone']]
    equipment_data = equipment_data[['name', 'start date', 'laptop', 'telefon sluzbowy', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
    content = f"Hi,<br>Please order a courier and prepare a delivery note for:<br>{shippment_data.to_html(index=False)}<br> Automatically generated email, addresse's was validated using google  address validate api."
    mail_sender.send_mail(user_id,"offce@lingarogroup.com", "Ordering a shipment courier", content, "2137", headers)
    if office_pick_up:
        office_pick_up = office_pick_up[['name', 'address','phone','Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
        mail_sender.send_mail(user_id,"sebastian.fraczak@lingarogroup.com", "SELF PICKUP", office_pick_up.to_html(index=False), "2137", headers)
    mail_sender.send_mail(user_id,"sebastian.fraczak@lingarogroup.com", "EQUIPMENT DATA", equipment_data.to_html(index=False), "2137", headers)

def check_email_tracker_list(employee_data, site_id, email_tracking_list_id, headers, mail_sender, shippment_data, office_pick_up, equipment_data, user_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{email_tracking_list_id}/items?expand=fields"
    response = requests.get(url=url, headers=headers)
    if response.status_code != 200:
        raise Exception(response.json())
    result = response.json()
    employee_details_to_add = []
    for employee in employee_data:
        current_date = datetime.now()
        employee_id = str(employee[2])
        already_on_list = False
        start_date = datetime.strptime(employee[3], "%Y-%m-%d %H:%M:%S")
        if start_date - timedelta(days=3) <= current_date <= start_date:  # when 3 days away from start date newcomer mail will be send.
            for data in result["value"]:
                fields = data["fields"]
                list_employee_id = fields['EmployeeId']
                if employee_id == list_employee_id:
                    shippment_data = shippment_data.drop(shippment_data[shippment_data['employeeID'].astype(str) == employee_id].index)
                    equipment_data = equipment_data.drop(equipment_data[equipment_data['employeeID'].astype(str) == employee_id].index)
                    already_on_list = True
                    break
            if not already_on_list:
                employee_details_to_add.append(employee) 
    if employee_details_to_add:
        add_sharepoint_email_tracking_record(site_id, email_tracking_list_id, headers, employee_details_to_add, mail_sender, shippment_data, office_pick_up, equipment_data,user_id)     

def main(logger,headers, application_id, drive_id, item_id, site_id, email_tracking_list_id, newbies_credentials_list_id, user_id):
    logger.info("Main init.")
    month_sheet, excel_sheets_data = get_excel_sheet(drive_id, item_id, headers)
    mail_sender = get_mail_sender_instance()
    sharepoint_data = get_sharepoint_data(headers, site_id, newbies_credentials_list_id)
    sulu_data = SuluData(application_id, headers)    
    for sheet in month_sheet:
        newcomers_excel_data, shippment_data,office_pickup_data, equipment_data = get_newcomers_data(drive_id, item_id, headers, sheet)
        print(newcomers_excel_data)
        employee_data = process_employee_data(sulu_data, newcomers_excel_data,sharepoint_data, mail_sender, headers, user_id)
        print(employee_data)  
        check_email_tracker_list(employee_data,site_id,email_tracking_list_id,headers, mail_sender, shippment_data, office_pickup_data, equipment_data, user_id)
    
if __name__ == "__main__":
    logger = setup_logger()
    load_dotenv("/Users/maciejcichocki/Documents/GitHub/newcomers_process_automation/Newcomers-Automation-Process/token.env")
    drive_id = os.getenv("DRIVE_ID")    #Sharepoint Data
    item_id = os.getenv("ITEM_ID")    #Sharepoint Data
    tenant_id = os.getenv("PYTHON_TENANT_ID")    # AZURE APP ID'S
    client_id = os.getenv("PYTHON_CLIENT_ID")    # AZURE APP ID'S
    client_secret = os.getenv("PYTHON_CLIENT_SECRET")    # AZURE APP ID'S
    application_id = os.getenv('APPLICATION_ID')    # AZURE APP ID'S
    username = os.getenv("USERNAME")    # AZURE APP ID'S
    site_id = os.getenv("SITE_ID")
    email_tracking_list_id = os.getenv("EMAIL_TRACKING_LIST_ID")
    send_grid_headers = os.getenv("SEND_GRID_CREDENTIALS")
    newbies_credentials_list_id = os.getenv("NEWBIES_CREDENTIALS_LIST_ID")# Newbies Credentials
    logger.info("Env variables Loaded.")
    mail_sender = MailSender()
    user_id = "maciej.cichocki@lingarogroup.com"
    headers = asyncio.run(msgraph_sdk_menager(tenant_id=tenant_id, client_id=client_id, client_secret=client_secret))
    logger.info("msgrapg_sdk connected.")
    main(
        logger = logger,
        headers=headers,
        application_id=application_id,
        drive_id=drive_id,
        item_id=item_id,
        site_id=site_id,
        email_tracking_list_id=email_tracking_list_id,
        newbies_credentials_list_id = newbies_credentials_list_id,
        user_id = user_id
    )