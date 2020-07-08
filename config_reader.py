import configparser
import sys, os, fnmatch
from datetime import datetime
# Prepare the spreadsheets to copy from and paste too.
config = configparser.ConfigParser()
#config.read(r'cfg.ini')
print(os.path.dirname(os.path.realpath(__file__)))
ini_file = os.path.dirname(os.path.realpath(__file__)) + os.path.sep + 'cfg.ini'
config.read(ini_file)
print("Script :" + sys.argv[0])
env = sys.argv[1]
myyear = sys.argv[3]
myper = sys.argv[2]
def get_config(env='PROD',myyear=datetime.today().strftime("%Y"),myper='01'):
    input_path = config.get(env, 'input_path')
    template_path = config.get(env, 'template_path')
    output_path = config.get(env, 'output_path')

    if env == 'DEV1' or env == 'DEV3':
        print("No parsing of period and year folders required based on " + env + " Environment parameter")
    elif env == 'DEV2' or env == 'QA' or env == 'PROD_TEST' or env == 'PROD':
        input_path = r'{}{}\Period {}\P{} {} Daily Reports\DAILYOUTPUT\\'.format(input_path,myyear,myper,myper,myyear)
    #    template_path = config[env]['template_path']
        output_path = r'{}{}\Period {}\P{} {} Daily Reports\DAILYOUTPUT\\'.format(output_path,myyear,myper,myper,myyear)
    else:
        print("Environment parameter is not matching")

    print("input path = " + input_path)
    return (input_path,template_path,output_path)