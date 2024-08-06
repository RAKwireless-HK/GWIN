from datetime import timedelta

class SIM_Info:
    def __init__(self, Project, SIM_ICCID, SIM_MRT, SIM_CARRIER, SIM_APN, SIM_IP, ACTIVE_Date):
        self.ORGANIZATION = "EMSD"
        self.PROJECT = Project
        self.SIM_ICCID = SIM_ICCID
        self.SIM_MRT = SIM_MRT # Mobile Phone number
        self.SIM_CARRIER = SIM_CARRIER # Mobile Phone Service Provider
        self.SIM_APN = SIM_APN # Access Point Name setting for the SIM card (APN)
        self.SIM_IP = SIM_IP
        self.ACTIVE_Date = ACTIVE_Date
        # self.END_Date = ACTIVE_Date + 365 days
        self.END_Date = ACTIVE_Date + timedelta(days=365)
        self.SIM_OWNER = "OWNER" # need to revise....

class GStack_Webpage:
    def __init__(self, _gstack):
        self.Domain = _gstack['Domain']
        self.User_Name = _gstack['User_Name']
        self.Password = _gstack['Password']

class EMSD_Project:
    def __init__(self, _emsd_project):
        self.organization = _emsd_project['Organization']
        self.project_code_01 = _emsd_project['Project_Code_01']
        self.project_code_02 = _emsd_project['Project_Code_02']
        self.project_code_03 = _emsd_project['Project_Code_03']
        self.project_Owner = _emsd_project['Project_Owner']

class GatewayInfo:
    def __init__(self, in_gateway_EUI, in_project_Code, in_owner, in_gateway_brand, in_gateway_model, in_antenna_type, in_antenna_dbi, in_Gateway_Firmware_version, in_VPN_KEY, in_SSH_USER, in_SSH_PASSWORD, in_Default_USER, in_Default_PASSWORD, in_Serial, in_Token):
        self.EUI = in_gateway_EUI.lower() if in_gateway_EUI is not None else ""
        self.Project_Code = in_project_Code
        self.Owner = in_owner
        self.Brand = in_gateway_brand
        self.Model = in_gateway_model
        self.Antenna_Type = in_antenna_type
        self.Antenna_DBI = in_antenna_dbi
        self.Firmware_Version = in_Gateway_Firmware_version
        self.VPN_KEY = in_VPN_KEY
        self.SSH_USER = in_SSH_USER
        self.SSH_PASSWORD = in_SSH_PASSWORD
        self.Default_USER = in_Default_USER
        self.Default_PASSWORD = in_Default_PASSWORD
        self.SERIAL = in_Serial
        self.Gateway_Token = in_Token
