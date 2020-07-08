import configparser
import sys, os, fnmatch
from datetime import datetime
# Prepare the spreadsheets to copy from and paste too.
config = configparser.ConfigParser()
#config.read(r'cfg.ini')
print(os.path.dirname(os.path.realpath(__file__)))
ini_file = os.path.dirname(os.path.realpath(__file__)) + os.path.sep + 'cfg.ini'
default_template_path = os.path.dirname(os.path.realpath(__file__)) + os.path.sep + 'templates' + os.path.sep
config.read(ini_file)

def get_config(env='PROD'):
    defaults_dic = {'year':datetime.today().strftime('%Y'),'per':'01','template_path':default_template_path}
    if env is None :
        env = 'PROD'
    input_path = config.get(env, 'input_path')
    template_path = config.get(env, 'template_path',fallback=defaults_dic['template_path'])
    output_path = config.get(env, 'output_path')
    myyear = config.get(env,'year',fallback=defaults_dic['year'])
    myper = config.get(env,'per',fallback=defaults_dic['per'])
#    myyear = year if year else defaults_dic['year']
#    myper = per if per else defaults_dic['per']
    if env == 'DEV1' or env == 'DEV3':
        print("No parsing of period and year folders required based on " + env + " Environment parameter")
    elif env == 'DEV2' or env == 'QA' or env == 'PROD_TEST' or env == 'PROD':
        input_path = r'{}{}\Period {}\P{} {} Daily Reports\DAILYOUTPUT\\'.format(input_path,myyear,myper,myper,myyear)
    #    template_path = config[env]['template_path']
        output_path = r'{}{}\Period {}\P{} {} Daily Reports\DAILYOUTPUT\\'.format(output_path,myyear,myper,myper,myyear)
    else:
        print("Environment parameter is not matching")

    print(input_path,template_path,output_path,myyear,myper)
    return (input_path,template_path,output_path,myyear,myper)