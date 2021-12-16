import os
from lxml import objectify
from SAPLogger.SapLogger import Logger


class SapLogon:
    def __init__(self) -> None:
        self.logger = Logger(log_name="SAPLogon")
        self.config_file = None
        self.xml_string = None
        self.xml_config = None
        self.remote_xml_config = None
        self.urls = []
        self.services_names = []
        self.get_config()

    def get_config(self) -> None:
        try:
            app_data = os.path.expandvars(r'%APPDATA%')
            if os.path.isfile(os.path.join(app_data, "SAP/Common/SAPUILandscape.xml")):
                self.config_file = os.path.abspath(os.path.join(app_data, "SAP/Common/SAPUILandscape.xml"))
            try:
                self.parse_xml_config()
            except Exception as err2:
                self.logger.log.error(err2)
            try:
                self.read_remote_config()
            except Exception as err2:
                self.logger.log.error(err2)
        except Exception as err:
            self.logger.log.error(err)
    
    def parse_xml_config(self) -> None:
        with open(self.config_file, "r") as f:
            self.xml_string = f.read()
            self.xml_config = objectify.fromstring(self.xml_string)
            self.urls.append(self.xml_config.Includes.Include.attrib['url'])
    
    def read_remote_config(self) -> None:
        for url in self.urls:
            url = url.split(":")[1]
            with open(url, "rb") as f:
                _xml_string = f.read()
                self.remote_xml_config = objectify.fromstring(_xml_string)
                for service in self.remote_xml_config.Services.getchildren():
                    self.services_names.append(service.attrib['name'])
    
    def sap_logon_pad(self) -> str:
        import PySimpleGUI as sg
        layout = [[sg.Listbox(values=self.services_names, size=(30, 6), enable_events=True, bind_return_key=True)]]
        window = sg.Window('Select SAP System', layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            else:
                window.close()
                if len(values) == 1:
                    v = values[0]
                    if len(v) == 1:
                        return v[0]
                    else:
                        return v
                else:
                    return values
