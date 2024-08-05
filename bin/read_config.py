# ReadConfig -------------------------------------------------------------------------------------------
class ReadConfig(object):
    """ This class parses configuration parameters from the configuration file """
# ------------------------------------------------------------------------------------------------------
    def __init__(self, config_file):
        
        config_kvp = {}

        try:

            with open(config_file) as cfp:
                for cl, config_line in enumerate(cfp):
                
                    kvp_string = config_line.strip()
                    if len(kvp_string) == 0 or kvp_string.find("#") > -1: # ignore empty lines and comments in config file
                        continue
                
                    kvp = kvp_string.split("=")
                    config_kvp[kvp[0].strip()] = kvp[1].strip()
        except:
            print("### FATAL ERROR: Unable to open or read the configuration file: {}").format(config_file)
            exit(-1)
        else:
            self.service_user        = config_kvp.get('service_user','')
            self.credentials_file    = config_kvp.get('credentials_file','')
            self.secret              = config_kvp.get('secret','')
            self.token               = config_kvp.get('token','')
            self.TBconfig_file       = config_kvp.get('TBconfig_file','')
            self.data_file           = config_kvp.get('data_file','')
            self.sheet_name          = config_kvp.get('sheet_name','')
            self.script_id           = config_kvp.get('script_id','')
            self.script_scope        = config_kvp.get('script_scope','')
            self.MAX_RETRY_COUNT     = int(config_kvp.get('MAX_RETRY_COUNT',2))
            self.MAX_CHEST           = int(config_kvp.get('MAX_CHEST',20))
            self.mailTo              = config_kvp.get('mailTo','')
            self.SUBJECT             = config_kvp.get('SUBJECT','')
            self.BROWSER             = config_kvp.get('BROWSER','Chrome')
            

            
            
            
            
