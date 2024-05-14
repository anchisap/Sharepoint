from shareplum import Site, office365
from shareplum.site import Version
import json
import os

# ปรับเส้นทางของไฟล์ config ให้ทำงานได้กับทุกแพลตฟอร์ม
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(ROOT_DIR, 'config.json')

# อ่านไฟล์ config.json
with open(config_path, 'r') as config_file:
    config = json.load(config_file)  # แก้ไขการพิมพ์ผิดจาก cofig_file เป็น config_file
    config = config['share_point']

USERNAME = config['user']
PASSWORD = config['password']
SHAREPOINT_URL = config['url']
SHAREPOINT_SITE = config['site']

class SharePoint:
    def auth(self):
        self.authcookie = office365(
            SHAREPOINT_URL,
            username=USERNAME,
            password=PASSWORD,
        ).GetCookies()
        self.site = Site(
            SHAREPOINT_SITE,
            version=Version.v365,  # แก้ไขการเขียน version ให้ถูกต้องตามการใช้งานของ shareplum
            authcookie=self.authcookie,
        )
        return self.site
    
    def connect_to_list(self, list_name):  # เปลี่ยนจาก is_name เป็น list_name
        self.auth_site = self.auth()
        list_data = self.auth_site.List(list_name=list_name).GetListItems()
        return list_data

# ตัวอย่างการใช้งาน:
# sp = SharePoint()
# data = sp.connect_to_list('ชื่อของลิสต์')
# print(data)
