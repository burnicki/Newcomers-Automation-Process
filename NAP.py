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
import sys


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

        employee_data = []
        for item in data['value']:
            fields = item['fields']
            employee_id = fields['Title']
            entra_id = fields['AzADObjectId']
            onepassword_link = fields['PasswordShareLink']
            employee_data.append([employee_id,entra_id,onepassword_link])
        return employee_data
    
    def get_sharepoint_email_tracker(self, site_id, email_tracker_list_id):
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{email_tracking_list_id}/items?expand=fields"
        response = requests.get(url=url, headers=headers)
        if response.status_code != 200:
            raise Exception(response.json())
        result = response.json()
        return result          
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
        formatted_address = result['result']['address']['formattedAddress']  
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
        logger.debug(f"Raw Dataframe from create_datafeame method: \n{df.head()}")
        return df
    
    def filter_df(self,raw_df):
        logger.info("Filtering data.")
        if raw_df.empty:
            raise Exception(logger.error("Empty dataframe was passed to filter_df"))
        data = ['employeeID', 'name', 'address', 'phone', 'start date', 'e-mail before start', 'laptop','telefon sluzbowy', 'umowa', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']
        filt = (~raw_df['address'].str.contains("Mexico|MEXICO|México", na=False))
        filtered_df = raw_df.loc[filt, data]
        return filtered_df
    
    def drop_missing_values(self, raw_df):
        logger.info("Droping NAN values")
        raw_df = raw_df.dropna(subset = ['name'])
        raw_df = raw_df.dropna(subset=['address'])
        raw_df = raw_df.dropna(subset=['employeeID'])
        raw_df.drop(raw_df[raw_df['umowa'] != "podpisana"].index, inplace = True)
        return raw_df
    
    def unidecode_name(self, raw_df):
        logger.info("Removing all special characters, lowercasing, spaces")
        raw_df['name'] = raw_df['name'].apply(unidecode).str.strip().str.lower()
        return raw_df
    
    def update_values(self,raw_df):
        logger.info("Updating values.")
        raw_df['laptop'] = raw_df['laptop'].replace(np.nan, "standard win" ,regex = True)
        raw_df['employeeID'] = raw_df['employeeID'].astype(int)
        raw_df['telefon sluzbowy'] = raw_df['telefon sluzbowy'].replace(np.nan, " " ,regex = True)
        raw_df['phone'] = raw_df['phone'].astype(str).str.replace(" ", "")
        raw_df['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'] = raw_df['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'].replace(np.nan, " ", regex = True)
        return raw_df
    
    def clean_newcomers_excel_data(self, raw_df):
        logger.info("Method clean_newcomers_excel_data init.")
        filtered_df = self.filter_df(raw_df)
        df_no_missing_values = self.drop_missing_values(filtered_df)
        df_unidecoded_names = self.unidecode_name(df_no_missing_values)
        processed_df = self.update_values(df_unidecoded_names)
        logger.info("Data cleaning complete.")
        logger.debug(processed_df.to_string())
        return processed_df

    def calculate_days_to_start(self, processed_df):
        logger.info(f"Method calculate_days_to_start init.")
        current_date = datetime.today()
        indexes = []
        if processed_df.empty:
            logger.error("Error, empty dataframe was passed to calculate_days_to_start")
            return pd.DataFrame()
        for index, row in processed_df.iterrows():
            start_date = row['start date'] 
            logger.info(f"{row['name']} | {start_date} | {type(start_date)}")
            if not isinstance(start_date, datetime):
                start_date = datetime.strptime(start_date, '%d.%m.%Y')
                logger.warning(f"Wrong start date format was found in: \n {row}")
            if current_date <= start_date <= current_date + timedelta(days=5) and start_date.weekday() in [0, 1, 5, 6]:  
                indexes.append(int(index))  
            elif start_date - timedelta(days=3) <= current_date <= start_date and start_date.weekday() in [2, 3, 4]: 
                indexes.append(int(index))
        return indexes
    
    def df_address_validation(self,processed_df ,indexes):
        if not indexes:
            logger.error("No indexes match conditions on calculate_days_to_start. Program will exit.")
            sys.exit(1)
        validad_addresses = []
        df = processed_df.loc[indexes]
        for address in df['address'].values.tolist():
            validate = self.validate_address(self.key,address)
            validad_addresses.append(validate)
        df['address'] = validad_addresses
        logger.info("Dataframe was updated with valid addresses. ")
        logger.debug(df.to_string())
        return df
    
    def extract_df_data(self,df):
        self_pickup = None
        equpiment_data = df[['employeeID','name', 'start date', 'laptop', 'telefon sluzbowy', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
        if not df['Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)'].str.strip().str.lower().eq("osobiście odbiór".strip().lower()).any():
            shippment_data = df[['employeeID','name', 'address', 'phone']]
        else:
            self_pickup = df[['employeeID','name', 'address','phone','Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]

        return equpiment_data, shippment_data, self_pickup

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
                            'address': address
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
        employee_start_date = employee_start_date_with_time.strftime("%d %B %Y %H:%M")
        employee_onepassword_active_date = datetime.combine(
            employee_start_date_with_time - timedelta(days=2),
            time(hour=4, minute=0)
        )
        credentials_enabled_from = employee_onepassword_active_date.strftime("%d %B %Y %H:%M")
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
                "FirstDayOfWork": employee_start_date,
                "OfficeCountryCode": "PL",
                "OnePasswordUrl": onepassword_link,
                "FirstName": employee_name[0],
                "AccountEnabledFrom": credentials_enabled_from
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
        self.raw_df = self.newcomers.create_dataframe(sheet)
        self.cleaned_df = self.newcomers.clean_newcomers_excel_data(self.raw_df)
        self.indexes = self.newcomers.calculate_days_to_start(self.cleaned_df)
        self.processed_df = self.newcomers.df_address_validation(self.cleaned_df,self.indexes)
    
    def get_excel_data(self):
        return self.processed_df
    
    def extract_shipping_data(self):
        logger.info("Extracting all shipping data")
        equipment_data, shippment_data, self_pickup = self.newcomers.extract_df_data(self.processed_df)
        return equipment_data, shippment_data, self_pickup
      
class SharepointMenager():
    def __init__(self,headers, site_id, newbies_credentials_list_id, email_tracking_list_id):
        self.sharepoint = SharepointData(headers)
        self.site_id = site_id
        self.newbies_credentials_list_id = newbies_credentials_list_id
        self.email_tracking_list_id = email_tracking_list_id
        
    def get_newbies_credentials(self):
        self.newbies_credentials = self.sharepoint.get_sharepoint_newbies_credentials(site_id, newbies_credentials_list_id)
        return self.newbies_credentials
    
    def get_email_tracking_list(self):
        self.email_tracking_list = self.sharepoint.get_sharepoint_email_tracker(site_id, email_tracking_list_id)
        return self.email_tracking_list 
    
    
        
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
    logger = logging.getLogger(my_logger) 
    return logger     
      
def process_string(s):
    s = unidecode(s)
    s = s.strip()
    s = s.lower()
    return s

def get_mail_sender_instance():
    return MailSender()    

def get_sharepoint_newcomers_credentials(headers, site_id, newbies_credentials_list_id,email_tracking_list_id):
    menager = SharepointMenager(headers, site_id, newbies_credentials_list_id,email_tracking_list_id)
    newcomers_credentials = menager.get_newbies_credentials()
    email_tracking_list = menager.get_email_tracking_list()
    return newcomers_credentials, email_tracking_list

def get_extract_email_tracking_employee_id(data):
    list_employee_id = [x["fields"]["EmployeeId"] for x in data["value"]]
    return list_employee_id

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
    process_df = menager.get_excel_data()
    equipment_data, shippment_data, self_pickup = menager.extract_shipping_data()
    return process_df, equipment_data, shippment_data, self_pickup

def prepare_employee_data(sulu_data, newcomers_excel_data, sharepoint_data):
    logger.info("Starting process_employee_data.")

    employee_data_for_sharepoint_email_tracking_list = set()
    employes_from_ltl = set()
    
    for _, row in newcomers_excel_data.iterrows():
        excel_employee_id = row["employeeID"]
        excel_employee_name = row["name"]
        excel_employee_start_date = str(row["start date"])
        excel_employee_personal_mail = row["e-mail before start"]

        logger.info(f"Processing employee ID: {excel_employee_id}, Name: {excel_employee_name}")
        
        try:
            sulu_employee_data = sulu_data.get_sulu_data(excel_employee_id)
            sulu_employee_name = sulu_employee_data['displayName']
            logger.info(f"Retrieved Sulu data for employee ID: {excel_employee_id}, Name: {sulu_employee_name}")
            
            if excel_employee_name == process_string(sulu_employee_name):
                sulu_employee_ms_id = sulu_employee_data["id"]
                
                sharepoint_entry = next((sp for sp in sharepoint_data if int(sp[0]) == excel_employee_id), None)
                
                if sharepoint_entry:
                    logger.info(f"Found matching SharePoint record for employee ID: {excel_employee_id}")
                    onepassword_link = sharepoint_entry[2]
                else:
                    logger.info(f"No SharePoint record found for employee ID: {excel_employee_id}. Adding a default record.")
                    onepassword_link = "https://share.1password.com/s#ASz7Tdy5k5aHrNjBErHSltiy8VKcS8bSr9Udcc-Hcug"
                    employes_from_ltl.add((
                        excel_employee_id,
                        sulu_employee_ms_id,
                        sulu_employee_name,
                        excel_employee_start_date
                    ))
                
                employee_data_for_sharepoint_email_tracking_list.add((
                    sulu_employee_ms_id,
                    sulu_employee_name,
                    excel_employee_id,
                    excel_employee_start_date,
                    excel_employee_personal_mail,
                    onepassword_link,
                ))
                
        except AttributeError as e:
            logger.error(f"AttributeError processing employee ID: {excel_employee_id}, Error: {e}")
    
    df = pd.DataFrame(employes_from_ltl, columns=['Employee id', 'Microsoft id', 'Name', 'Start date']) if employes_from_ltl else None
    
    logger.info(f"Finished processing employees. Total employees for SharePoint tracking: {len(employee_data_for_sharepoint_email_tracking_list)}")
    return employee_data_for_sharepoint_email_tracking_list, df


def add_email_tracking_record_to_sharepoint(site_id, email_tracking_list_id, headers, employee_data):
    logger.info("Starting add_sharepoint_email_tracking_record.")
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{email_tracking_list_id}/items"
    
    for data in employee_data:
        entra_id = data[0]
        employee_name = data[1]
        employee_id = str(data[2])
        employee_start_date = data[3]  # format = '2024-07-22 00:00:00'        
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
        
        logger.info(f"Posting data to SharePoint for employee ID: {employee_id}")
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code != 201:
            logger.error(f"Failed to post data for employee ID: {employee_id}. Response: {response.json()}")
            raise Exception(response.json())
        
        logger.debug(f"Successfully posted data for employee ID: {employee_id}. Response: {json.dumps(response.json(), indent=2)}")


def send_welcome_emails(employee_data, mail_sender):
    logger.info("Starting send_emails.")
    
    for data in employee_data:
        employee_name = data[1]
        employee_personal_mail = data[4]
        employee_one_password = data[5]
        start_date = datetime.strptime(data[3], "%Y-%m-%d %H:%M:%S")
        
        logger.info(f"Sending welcome email to employee: {employee_name}")
        mail_sender.send_welcome_mail_to_newcomer(employee_name, start_date, employee_one_password, employee_personal_mail)

def send_logistics_emails(mail_sender, shipment_data, equipment_data, office_pickup_data,ltl_data, user_id):
    content = f"Hi,<br>Please order a courier and prepare a delivery note for:<br>{shipment_data.to_html(index=False)}<br> Automatically generated email, addresse's was validated using google address validate API."
    logger.info("Sending courier order email.")
    mail_sender.send_mail(user_id, "office@lingarogroup.com", "Ordering a shipment courier", content, "2137", headers)
    
    logger.info("Sending equipment data email.")
    mail_sender.send_mail(user_id, "sebastian.fraczak@lingarogroup.com", "EQUIPMENT DATA", equipment_data.to_html(index=False), "2137", headers)
    
    if office_pickup_data:
        logger.info("Sending self pickup email.")
        mail_sender.send_mail(user_id, "sebastian.fraczak@lingarogroup.com", "SELF PICKUP", office_pickup_data.to_html(index=False), "2137", headers)
    if ltl_data:
        content_ltl = f"Employee password must be reset to - L1n99aROrba22 <br> {ltl_data.to_html(index=False)}"
        logger.info(f"Sending email for long term leavers with {len(ltl_data)} employees.")
        mail_sender.send_mail(user_id, "dominik.boras@lingarogroup.com", "Long Term Leavers", content_ltl, "2137", headers)
         
    


def filter_newcomer_records_for_sharepoint(employee_data, sharepoint_employee_id):
    logger.info("Starting filter_newcomers_sharepoint_record.")
    current_time = datetime.now()
    employee_to_add = []
    
    for employee in employee_data:
        on_list = False
        employee_id = str(employee[2])
        start_date = datetime.strptime(employee[3], "%Y-%m-%d %H:%M:%S")
        
        if start_date - timedelta(days=3) <= current_time <= start_date:
            logger.info(f"Checking if employee ID: {employee_id} is already on SharePoint list.")
            for sharepoint_id in sharepoint_employee_id:
                if employee_id == sharepoint_id:
                    logger.info(f"Employee ID: {employee_id} found on SharePoint list.")
                    on_list = True
                    break
            
            if not on_list:
                logger.info(f"Employee ID: {employee_id} not found on SharePoint list. Adding to list.")
                employee_to_add.append(employee)
    
    logger.info(f"Filtered employees to add to SharePoint: {len(employee_to_add)}")
    return employee_to_add



def prepare_newcomer_shipping_data(employee_data, sharepoint_data, shippment_data, equipment_data, office_pickup_data):
    logger.info("Starting process_newcomers_shipping_data.")
    
    if not employee_data:
        logger.warning("Empty list of employee data.")

    
    logger.info("Cleaning up shipment and equipment data.")
 
    maching_ids = [x[2] for x in employee_data if x[2] in sharepoint_data]
    logger.warning(f"Employee already on sharepoint : {maching_ids}")   
    if maching_ids:
        shippment_data = shippment_data[~shippment_data['employeeID'].isin(maching_ids)]
        equipment_data = equipment_data[~equipment_data['employeeID'].isin(maching_ids)]
    
    shippment_data = shippment_data[['name', 'address', 'phone']]
    equipment_data = equipment_data[['name', 'start date', 'laptop', 'telefon sluzbowy', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
    if office_pickup_data != None:
        office_pickup_data = office_pickup_data[['name', 'address', 'phone', 'Dodatkowe( wczesniejsza wysylka lub odbiór osobisty)']]
    
    logger.info("Finished processing shipping and equipment data.")
    return equipment_data, shippment_data, office_pickup_data


def main(logger, headers, application_id, drive_id, item_id, site_id, email_tracking_list_id, newbies_credentials_list_id, user_id):
    logger.info("Main function init.")
    
    month_sheet, excel_sheets_data = get_excel_sheet(drive_id, item_id, headers)
    logger.warning(month_sheet)
    mail_sender = get_mail_sender_instance()
    newcomers_credentials, email_tracking_list = get_sharepoint_newcomers_credentials(headers, site_id, newbies_credentials_list_id, email_tracking_list_id)
    sharepoint_employee_id = get_extract_email_tracking_employee_id(email_tracking_list)
    
    logger.info(f"SharePoint employee IDs extracted: {len(sharepoint_employee_id)}")
    
    sulu_data = SuluData(application_id, headers)
    
    for sheet in month_sheet:
        logger.info(f"Processing sheet: {sheet}")
        newcomers_excel_data, equipment_data, shippment_data, office_pickup_data = get_newcomers_data(drive_id, item_id, headers, sheet)
        
        employee_data, ltl_data = prepare_employee_data(sulu_data, newcomers_excel_data, newcomers_credentials)
        filtered_employees = filter_newcomer_records_for_sharepoint(employee_data, sharepoint_employee_id)
        
        logger.info(f"Filtered employees: {len(filtered_employees)}")
        
        equipment_clean, shippment_clean, office_pickup_clean = prepare_newcomer_shipping_data(filtered_employees,sharepoint_employee_id ,shippment_data, equipment_data, office_pickup_data)
        add_email_tracking_record_to_sharepoint(site_id, email_tracking_list_id, headers, filtered_employees)
        # send_welcome_emails(filtered_employees, mail_sender)
        # send_logistics_emails(mail_sender,shippment_clean,equipment_clean,office_pickup_clean,ltl_data, user_id)
    
    logger.info("Main function completed.")
    
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