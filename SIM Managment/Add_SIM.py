import configparser
import time
import os
import re
import Web_Controller as wc
import Excel_Controller as ec
import Common_Obj as co

class GStack_Web_controller:
    def __init__(self, webpage: co.GStack_Webpage, _sleep_time):
        self.web_controller = wc.Web_Contrller(webpage, _sleep_time)

    def login(self):
        # 'username' 'password' 'login-button' login
        self.web_controller.login('username', 'password', 'login-button') 

    def goto_SIM_page(self):
        # '//*[@id="nav-bar-button"]' click
        self.web_controller.wait_element_load_n_click('//*[@id="nav-bar-button"]')
        # '//*[@id="left-menu-sim"]' click
        self.web_controller.wait_element_load_n_click('//*[@id="left-menu-sim"]')
    
    def goto_Add_SIM_page(self):
        # //*[@id="segment-box"]/div/div/a[1]/button
        self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div/a[1]/button')

    def set_organization(self, organization: str):
        # //*[@id="organization-select-dropdown"]/i Click
        self.web_controller.wait_element_load_n_click('//*[@id="organization-select-dropdown"]/i')
        if organization == "EMSD":
            # /html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div[2]/form[1]/div/div[1]/div/div[2]/div[2]
            self.web_controller.wait_element_load_n_click('/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div[2]/form[1]/div/div[1]/div/div[2]/div[2]')
        elif organization == "ATAL":
            #//*[@id="organization-select-dropdown"]/div[2]/div[1] click
            self.web_controller.wait_element_load_n_click('/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div[2]/form[1]/div/div[1]/div/div[2]/div[1]')

    def set_project(self, project: str):
        # //*[@id="project-select-dropdown"]/i Click
        self.web_controller.wait_element_load_n_click('//*[@id="project-select-dropdown"]/i')
        if project == "1696EM23M":
            # //*[@id="project-select-dropdown"]/div[2]/div[1]
            self.web_controller.wait_element_load_n_click('//*[@id="project-select-dropdown"]/div[2]/div[1]')
        elif project == "1697EM23M":
            # //*[@id="project-select-dropdown"]/div[2]/div[2]
            self.web_controller.wait_element_load_n_click('//*[@id="project-select-dropdown"]/div[2]/div[2]')
        elif project == "1698EM23M":
            # //*[@id="project-select-dropdown"]/div[2]/div[3]
            self.web_controller.wait_element_load_n_click('//*[@id="project-select-dropdown"]/div[2]/div[3]')

    def set_SIM_ICCID(self, SIM_ICCID: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[1]/div[1]/div/input  clear all information in the box
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[1]/div[1]/div/input', SIM_ICCID)

    def set_MRT(self, SIM_MRT: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[1]/div/input clear all information in the box
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[1]/div/input', SIM_MRT)

    def set_Carrier(self, SIM_CARRIER: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/i click
        self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/i')
        if SIM_CARRIER == "CSL":
            # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[1]
            self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[1]')
        elif SIM_CARRIER == "SmarTone":
            # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[2]
            self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[2]')
        elif SIM_CARRIER == "CMHK":
            # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[3]
            self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[3]')
        elif SIM_CARRIER == "3HK":
            # //*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[4]
            self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[3]/div[2]/div/div[2]/div[4]')

    def set_APN(self, SIM_APN: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[4]/div[1]/div/input clear all information in the box
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[4]/div[1]/div/input', SIM_APN)

    def set_Static_IP(self, SIM_IP: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[1]/div/i click
        self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[1]/div/i')
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[1]/div/div[2]/div[1] click
        self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[1]/div/div[2]/div[1]')

    def set_IP(self, IP_ADDRESS: str):
        # only take the IP in the IP_ADDRESS
        ip_pattern = r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b'
        ipaddress = re.findall(ip_pattern, IP_ADDRESS)
        IP_ADDRESS = ipaddress[0]                
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[2]/div/input clear all information in the box
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[5]/div[2]/div/input', IP_ADDRESS)

    def set_Active_Date(self, ACTIVE_Date: str):
        # ACTIVE_Date convert to YYYY-MM-DD fomat string
        _temp_date = ACTIVE_Date.strftime('%Y-%m-%d')
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[6]/div[1]/div/div/div/input
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[6]/div[1]/div/div/div/input', _temp_date)

    def set_End_Date(self, END_Date: str):
        _temp_date = END_Date.strftime('%Y-%m-%d')
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[6]/div[2]/div/div/div/input
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[6]/div[2]/div/div/div/input', _temp_date)

    def set_Owner(self, SIM_OWNER: str):
        # //*[@id="segment-box"]/div/div[2]/form[2]/div[7]/div[1]/div/input clear all information in the box
        self.web_controller.fillin_value('//*[@id="segment-box"]/div/div[2]/form[2]/div[7]/div[1]/div/input', SIM_OWNER)

    def click_Add_button(self):
        # //*[@id="segment-box"]/div/div[2]/form[2]/button click
        self.web_controller.wait_element_load_n_click('//*[@id="segment-box"]/div/div[2]/form[2]/button')        
        # /html/body/div[3]/div/div[3]/button click
        self.web_controller.wait_element_load_n_click('/html/body/div[3]/div/div[3]/button')
    
    def add_SIM(self, SIM_Info: co.SIM_Info):
        self.set_organization(SIM_Info.ORGANIZATION)        
        self.set_project(SIM_Info.PROJECT)
        self.set_SIM_ICCID(SIM_Info.SIM_ICCID)
        self.set_MRT(SIM_Info.SIM_MRT)        
        self.set_Carrier(SIM_Info.SIM_CARRIER)
        self.set_APN(SIM_Info.SIM_APN)        
        self.set_Static_IP(SIM_Info.SIM_IP)
        self.set_IP(SIM_Info.SIM_IP)
        self.set_Active_Date(SIM_Info.ACTIVE_Date)
        self.set_End_Date(SIM_Info.END_Date)
        self.set_Owner(SIM_Info.SIM_OWNER)
        #self.click_Add_button()

def read_SIM_info_from_Excel(in_file_path: str, page_name):
    excel_controller = ec.Excel_Controller(in_file_path)
    gatewayInfos = excel_controller.read_excel_and_convert_to_gateway_info(page_name)
    return gatewayInfos

# 讀取配置文件
config = configparser.ConfigParser()
#current path of the file
_current_path = os.path.dirname(os.path.realpath(__file__))
_setting_file = os.path.join(_current_path, 'config.ini')
config.read(_setting_file)
_excel_file_name = config['SIM_DATA_FILE']['Path']
_excel_page = config['SIM_DATA_FILE']['Page']
_sleep_time = config['Web_Controller']['Sleep_Time']
_gstack = {
    "Domain": config.get('Webpage', 'Domain'),
    "User_Name": config.get('Webpage', 'User_Name'),
    "Password": config.get('Webpage', 'Password')
}

_excel_path = os.path.join(_current_path, _excel_file_name)
print (_excel_path)

if not os.path.exists(_excel_path):
    print("File not found")

SIM_Card_Infos = read_SIM_info_from_Excel(_excel_path, _excel_page)

gStackInfo = co.GStack_Webpage(_gstack)

web_controller = GStack_Web_controller(gStackInfo,_sleep_time)
web_controller.login()
web_controller.goto_SIM_page()
web_controller.goto_Add_SIM_page()

for SIM_Card_Info in SIM_Card_Infos:
    web_controller.add_SIM(SIM_Card_Info)
    time.sleep(4)
    print("Below information is added to GStack:")
    print( 
            SIM_Card_Info.ORGANIZATION,
            SIM_Card_Info.PROJECT,
            SIM_Card_Info.SIM_ICCID, 
            SIM_Card_Info.SIM_MRT, 
            SIM_Card_Info.SIM_CARRIER, 
            SIM_Card_Info.SIM_APN, 
            SIM_Card_Info.SIM_IP, 
            SIM_Card_Info.ACTIVE_Date,
            SIM_Card_Info.END_Date,
            SIM_Card_Info.SIM_OWNER
          )
    time.sleep(4)
    web_controller.goto_Add_SIM_page()
    time.sleep(3)
    
print("All SIM Card information is added to GStack successfully!")

