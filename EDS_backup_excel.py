import shutil
import time

path = '//Administrator/общая/ЭЦП/'
shutil.copyfile(path + 'я_Список ЭЦП.xlsm', path + 'я_Список ЭЦП.xlsm.bak')
time.sleep(30)
shutil.copyfile(path + 'я_Список ЭЦП.xlsm', 'D:\\я_Список ЭЦП.xlsm.bak')