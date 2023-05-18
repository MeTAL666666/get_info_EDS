from glob import glob
import os, sys, shutil, locale
import json
import time
import pysftp
import paramiko
from datetime import datetime

locale.setlocale(
    category=locale.LC_ALL,
    locale="Russian"
)

connect = json.load(open('D:\\EDS_connect.json'))	# файл с параметрами подключения к SFTP
username = connect['username']						# 
password = connect['password']
path = connect['path']

list_not_sign = []
eds = {}

if os.path.exists(path + 'sftp_not_sign.txt'):
	os.remove(path + 'sftp_not_sign.txt')

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None  # disable host key checking.
# Connection to external SFTP server
with pysftp.Connection(host='103.58.1.13',
                       port=22,
                       username=username,
                       password=password,
                       cnopts=cnopts
                       ) as sftp:
	print('Connected to SFTP')
	with sftp.cd('/var/opt/cprocsp/keys/signegisz/'):
		print('Поиск контейнеров на сервере...',end='\r')
		files = sftp.listdir('./')
		for folder in files:
			if folder.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
				try:
					if sftp.stat('./' + folder + '/header.key').st_size < 1500:
						list_not_sign.append(folder)
				except:
					list_not_sign.append(folder)

if len(list_not_sign) > 0:
	with open(path + 'sftp_not_sign.txt','a')as file:
		file.write(datetime.strftime(datetime.now(),'%d.%m.%y %a %H:%M') + '\n' )
		for i, name_eds in enumerate(list_not_sign):
			print(f'Найдено неподписанных контейнеров: {i + 1}')
			print('Идёт сопоставление ФИО...', end='\r')
			search_name = glob(path + '**/%s' % name_eds, recursive=False)
			if len(search_name) != 0:
				search_name = search_name[-1].split('\\')[-2]
				file.write(search_name + '\n')
			elif len(search_name) == 0:
				file.write(name_eds + '\n')
	cnt = 3
	while cnt > 0:
		cnt -= 1
		time.sleep(1)
		print(f'Отображение результата...({cnt})', end='\r')
	os.chdir(path)
	os.startfile (r"sftp_not_sign.txt")
	print('Открыт файл с результатами          ', end='\r')
	time.sleep(2)
	#os.system(f"notepad.exe {path}sftp_not_sign.txt") # программа не будет выполняться дальше, пока открыт файл
else:
	print('Неподписанных контейнеров на сервере на найдено')
	time.sleep(2)