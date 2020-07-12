
from config_reader import *
import datetime
(TB_input_path,PL_input_path,template_path,TB_output_path,PL_output_path,myyear,myper) = get_config(env=sys.argv[1] if len(sys.argv) > 1 else None)

print(os.getcwd())

current_date_and_time = datetime.datetime.now()
print(current_date_and_time)

#buffer to handle post process time
#hours = 2
hours_substracted = datetime.timedelta(hours = 2)
updated_date_time = current_date_and_time - hours_substracted
print(updated_date_time)

am_pm = updated_date_time.strftime("%p").lower()
month = (updated_date_time.strftime("%m"))
day = (updated_date_time.strftime("%d"))
cal_year = (updated_date_time.strftime("%y"))

new_folder_name = '{}.{}.{}_11.00{}'.format(month, day, cal_year, am_pm)
print(new_folder_name)

#folder_path = r'{}'.format(TB_input_path)
print(TB_input_path)
os.chdir(TB_input_path)
print(os.getcwd())
#os.chdir("../../")
os.chdir('..{}..{}'.format(os.path.sep,os.path.sep))
print(os.getcwd())

os.rename("DAILYOUTPUT",new_folder_name)

