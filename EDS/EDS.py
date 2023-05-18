from pathlib import Path
from glob import glob
from ctypes import windll
from datetime import datetime, timedelta
import os, sys, re, shutil, time
import zipfile
import json
import signal

#from inputimeout import inputimeout, TimeoutOccurred
import pysftp
import paramiko
import win32api, win32con


# Функция для правильного форматирования слова 'дней'
def days_remained(days):
	if int(days) >= 0:
		remained = int(days) % 10
		if days == 0:
			return ' Истекает СЕГОДНЯ!'
		elif remained == 1 and days !=11:
			return days + ' день'
		elif remained == 0 or days in range(11,20):
			return days + ' дней'
		elif days in range (2,5):
			return days + ' дня'
		elif remained in range(5,10):
			return days + ' дней'
		elif remained in range(2,5):
			return days + ' дня'
			
	elif int(days) < 0:
		days = days[1:]
		remained = int(days) % 10
		if remained == 1 and days !=11:
			return 'Прекратил действовать ' + days + ' день назад'
		elif remained == 0 or days in range(11,20):
			return 'Прекратил действовать ' + days + ' дней назад'
		elif days in range (2,5):
			return 'Прекратил действовать ' + days + ' дня назад'
		elif remained in range(5,10):
			return 'Прекратил действовать ' + days + ' дней назад'
		elif remained in range(2,5):
			return 'Прекратил действовать ' + days + ' дня назад'
			
#Функция для определения периода действия сертификата
def period(name, cert_file):
	with open(path + name + '/' + cert_file,  encoding='UTF-8', errors="replace") as cert:
		a = cert.readlines()
		period = []
		for j in a:
			b = re.search(r"^2\d{11}Z", j.strip())
			if b != None:
				period.append([b.group()[0:6][i:i+2] for i in range(0, len(b.group()[0:6]), 2)])
	
	begin = period[0]
	end = period[1]

	begin.reverse()
	end.reverse()

	begin = datetime.strptime('.'.join(begin), '%d.%m.%y')
	end = datetime.strptime('.'.join(end), '%d.%m.%y')

	begin_str = begin.strftime('%d.%m.%Y')
	end_str = end.strftime('%d.%m.%Y')
	remained_str = str(end - datetime.now() + timedelta(days=1)).split(' ')[0]
	return {'begin': begin_str, 'end': end_str, 'remained': remained_str}

# Функция для очистки буфера обмена
def delFromClipboard():
	if windll.user32.OpenClipboard(None):
		windll.user32.EmptyClipboard()
		windll.user32.CloseClipboard()

# Функция для помещения в буфер обмена
def addToClipBoard(text):
    command = 'echo ' + text.strip() + '| clip'
    os.system(command)
# Функция для копирования файлов
def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)
# Функция для создания анимаций в надписях
def animation(anim_cycle, text ):
	cnt = 0
	anim = 0
	while anim < anim_cycle:
		while cnt < 5:
			cnt +=1 
			time.sleep(0.15)
			print(text + ' .' * cnt, end="\r")
		print(text + '  ' * 5,end="\r")
		cnt = 0
		anim += 1
		if anim == anim_cycle:
			cnt = 4
			print(text + ' .' * cnt, end="\r")

os.chdir('D:\\')									# смена директории ПО для правильной работы на разных ПК

connect = json.load(open('D:\\EDS_connect.json'))	# файл с параметрами подключения к SFTP и 
exceptions = {}										# словарь для загрузки нетипичных названий папок

# Параметры подключения к SFTP
username = connect['username']						
password = connect['password']

path = connect['path']								# путь к общей папке
backup_dir = 'backup_sent_EDS'						# имя папки с бэкапами переданных подписей

dirs_on_D = []										# список контейнеров на диске D:\\
list_dirs = []										# список папок с подписями(ФамилияИО)
eds = {}											# словарь с параметрами ЭЦП для каждого сотрудника

# Создание корневой папки с бэкапами
try:
	os.mkdir(path + backup_dir)
except:
	pass
# Чтение файла с исключительными названиями
with open ('EDS_exceptions.txt', 'r', encoding='cp1251') as file:
	for i in file:
		i = i.replace('\t', ' ').replace('ё','е').replace('Ё', 'Е').strip().split(':')
		exceptions[i[0]] = i[1]

# Сравнение подписанных контейнеров на D:\\ c логом отправленных			
try:
	with open('EDS_log.txt', 'r') as file:
		file_list = []
		list_eds = []
		for i in file:
			file_list.append(i.strip())
		for i in reversed(file_list):
			if i.startswith('-'):
				break
			elif i != '\n' and i != '':
				list_eds.append(i.strip())

		folders_sign = []							# Список с подписанными контейнерами на диске D:\\
		for folder in os.listdir(os.getcwd()):
			if os.path.isdir(folder):
				if folder.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
					# Если файл header.key < 1500 байт - он считается неподписанным
					if Path(os.getcwd() + folder +  '/header.key').stat().st_size < 1500:
						shutil.rmtree(folder)
					elif Path(os.getcwd() + folder +  '/header.key').stat().st_size >= 1500:
						if os.path.exists('EDS.tmp'):
							folders_sign.append(folder.strip())
						if not os.path.exists('EDS.tmp'):
							shutil.rmtree(folder)
		# Файл для отключения проверки подписанных контейнеров при первом запуске ПО
		if not os.path.exists('EDS.tmp'):
			with open ('EDS.tmp', 'w') as file:
				pass
			win32api.SetFileAttributes('EDS.tmp',win32con.FILE_ATTRIBUTE_HIDDEN)	# сделать файл скрытым
		
		folders_sign_to_del = []	# Список для удаления подписанных контейнеров при запуске ПО
		if len(folders_sign) > 0:
			for i in folders_sign:
				if i.strip() not in [i.split(' ')[1].strip() for i in list_eds]:
					folders_sign_to_del.append(i.strip())
				else:
					shutil.rmtree(i.strip())
		if len(folders_sign_to_del) > 0:
			print('Обнаружены подписанные, но непереданные ЭЦП!')
			animation(4,'Идёт поиск ФИО для контейнеров')
			
			for name_eds in folders_sign_to_del:
				search_name = glob(path + '**/%s' % name_eds, recursive=False)
				if len(search_name) != 0:
					search_name = search_name[-1].split('\\')[-2]
					eds[name_eds] = {'name': search_name, 'EDS': name_eds}
				elif len(search_name) == 0:
					eds[name_eds] = {'name': 'N/A', 'EDS': name_eds}
			print(' '*50,end='\r')
			print([i['name'] + ': ' +  i['EDS'] for i in eds.values()])

			while True:
				msg = input('Удалить контейнеры - (ввести 1) | Оставить контейнеры - (ввести 2)\n')
				if msg == '1':
					for i in folders_sign_to_del:
						shutil.rmtree(i)
						del eds[i]
					print('Удаление контейнеров...', end='\r')
					time.sleep(2)
					folders_sign_to_del.clear()
					print('Подписанные, но непереданные контейнеры УДАЛЕНЫ')
					break
				elif msg == '2':
					print('Подписанные, но непереданные контейнеры ОСТАВЛЕНЫ')
					break
				else:
					continue
except:
	for folder in os.listdir(os.getcwd()):
		if os.path.isdir(folder):
			if folder.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
					shutil.rmtree(folder)
# Чтение данных из файла EDS_input.txt
restart_input = 0	# переменная для повторного чтения файла данных
while True:					
	with open ('EDS_input.txt', 'r', encoding='cp1251') as file:
		not_use_input = 0	# переменная для меню выбора данных 
		for i in file:
			flag = 0
			if 	i in exceptions.values() and not i.startswith('\n'):
				list_dirs.append(i)
				flag = 1
				
			if i.strip()[-2:-1].islower() and flag != 1:
				i = i.replace('\t', ' ').replace('ё','е').replace('Ё', 'Е').strip()
				if i not in exceptions.keys() and not i.startswith('\n'):
					i = i.split(' ')
					i[0] = i[0].capitalize()
					i[1] = i[1][0]
					i[2] = i[2][0]
					i = ''. join(i)
					list_dirs.append(i)
				elif i not in exceptions.keys() and i.startswith('\n'):
					pass
				elif i in exceptions.values() and not i.startswith('\n'):
					list_dirs.append(i)
				else:
					i = exceptions[i]
					list_dirs.append(i)
			if i.strip()[-2:-1].isupper() and flag != 1:
				if not i.startswith('\n'):
					i = i.replace(' ','')
					i = i[:-2].capitalize() + i[-2:]
					list_dirs.append(i)
				elif i.startswith('\n'):
					pass

	if len(list_dirs) > 0:
		if restart_input == 0:
			print('Использовать данные из файла EDS_input.txt?')
			msg = input(f"Нет - (нажать Enter) | Да - (ввести 2) | Открыть файл -  (ввести 3) | "
								f"ОЧИСТИТЬ файл - (ввести 4)\n")#, timeout=5)
			#except TimeoutOccurred:
			if msg == '':
				not_use_input = 1
				list_dirs.clear()
				break
			elif msg == '2':
				break
			elif msg == '3':
				mtime = os.stat('EDS_input.txt').st_mtime
				for i in range(2):
					print('Программа продолжит работу после закрытия файла EDS_input.txt!', end='\r')
					time.sleep(0.3)
					print(' '*65, end='\r')
					time.sleep(0.3)
					print('Программа продолжит работу после закрытия файла EDS_input.txt!', end='\r')
				os.system(f"notepad.exe EDS_input.txt")
				print(' '*65, end='\r')
				if mtime != os.stat('EDS_input.txt').st_mtime:
					time.sleep(1)
					print('Файл данных успешно изменён!\n')
					time.sleep(1)
					restart_input = 1
					continue
				else:
					break
			elif msg =='4':
				os.remove('EDS_input.txt')
				list_dirs.clear()
				with open ('EDS_input.txt', 'w', encoding='cp1251') as file:
					pass
				print('Файл данных ОЧИЩЕН!\n')
				break
		elif restart_input == 1:
			break
	elif len(list_dirs) == 0:
		break
# Работа ПО с пустым EDS_input.txt
if len(list_dirs) == 0:
	if not_use_input != 1:
		print('Файл EDS_input.txt пуст.')
	print('Введите одно или несколько ФИО (разделяя запятой) в формате: '\
		'ФамилияИО или Фамилия Имя Отчество')
	while True:
		msg = input()
		try:
			msg = msg.split(',')
			for i in msg:
				flag = 0
				if len([j for j in i.strip() if j.isdigit() == True]) > 0:
					raise TypeError('В ФИО присутствуют цифры.')
				
				if i.strip() in exceptions.values() and not i.startswith('\n'):
					list_dirs.append(i.strip())
					flag = 1
					
				if i.strip()[-2:-1].islower() and flag != 1:
					i = i.replace('\t', ' ').replace('ё','е').replace('Ё', 'Е').strip()
					if i not in exceptions.keys():
						i = i.split(' ')
						i[0] = i[0].capitalize()
						i[1] = i[1][0]
						i[2] = i[2][0]
						i = ''. join(i)
						list_dirs.append(i)
					elif i not in exceptions.keys():
						pass
					else:
						i = exceptions[i]
						list_dirs.append(i)
				if i.strip()[-2:-1].isupper() and flag != 1:
					i = i.replace(' ','')
					i = i[:-2].capitalize() + i[-2:]
					list_dirs.append(i)
				flag = 0
			break
			print(list_dirs)
		except Exception as e:
			if type(e) == TypeError:
				print('Неверный формат ввода ФИО, повторите попытку. ' + str(e))
			else:
				print('Неверный формат ввода ФИО, повторите попытку.')
			list_dirs.clear()
			flag = 0
			continue
print()
for i in list_dirs:
	i = i.strip()
	try:
		list_files = os.listdir(path + i)
	except FileNotFoundError:
		print(f"Не найден путь: {path}{i}")
		os.system('pause')
		sys.exit()
	eds[i] = {'name':i,'EDS':'', 'password': ''}
	for j in list_files:
		if j.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
			eds[i]['EDS'] = j
		if j.endswith('.txt'):
			try:
				with open (path + i + '/' + j) as f:
					password_eds = re.findall(r'\n\d{6}',f.read())[-1].strip()
					eds[i]['password'] = password_eds
			except:
				eds[i]['password'] = ''	
# Запись лога найденных и переданных контейнеров
with open ('EDS_log.txt', 'w') as file:
	for key,value in eds.items():
		time.sleep(1.5)
		file.write(f"{value['name']}: {value['EDS']}\n")
		try:
			os.mkdir('D:\\' + value['EDS'])
			copytree(path + value['name'] + '/' + value['EDS'], 'D:\\' + value['EDS'])
			try:
				sent_EDS = os.listdir(f"{path}/{backup_dir}")
				if (f"{value['EDS']} {value['name']}") in [was_sent for was_sent in sent_EDS]:
					was_sent = datetime.strftime(datetime.fromtimestamp(os.path.getmtime(path + backup_dir +'/'
								+ (f"{value['EDS']} {value['name']}"))), '%d.%m.%y %H:%M')
					print(f"{value['name']}: {value['EDS']} --- {value['password']} [Уже был отправлен: {was_sent}]")
				else:
					print(f"{value['name']}: {value['EDS']} --- {value['password']}")
			except KeyError:
				pass
			dirs_on_D.append((value['name'], value['EDS']))
		except PermissionError:
			not_container = []
			try:
				with open('EDS_not_container.txt', 'r') as file_nc:
					for i in file_nc:
						not_container.append(i.strip())
			except:
				with open('EDS_not_container.txt', 'w') as f:
					pass
			with open('EDS_not_container.txt', 'a') as f:
				if value['name'] not in not_container:
					f.write(f"{value['name']}\n")
			print(f"{value['name']}: Папка НЕ СКОПИРОВАНА на диск D:\. Отсутствует контейнер")
			pass
		except FileExistsError:
			if value['EDS'] not in folders_sign_to_del:
				shutil.rmtree('D:\\' + value['EDS'])
				os.mkdir('D:\\' + value['EDS'])
				copytree(path + value['name'] + '/' + value['EDS'], 'D:\\' + value['EDS'])
				try:
					print(f"{value['name']}: {value['EDS']} --- {value['password']}")
				except KeyError:
					print(f"{value['name']}: {value['EDS']}")
				dirs_on_D.append((value['name'], value['EDS']))
			else:
				dirs_on_D.append((value['name'], value['EDS']))
	file.write('\n' + '-' * 20 + '\n')
print('\nУстановите сертификаты пользователей:')
time.sleep(1)

temp = []
for m in eds.items():
	if re.search(r'\d{3}$', m[0]) is None: 
		temp.append(m)

if len(temp) > 0:
	for n,j in enumerate(temp):
		while True:
			cert_current = 'сертификат'
			pass_current = 'пароль от подписи'
			container_current = 'контейнер'
			zip_current = 'zip-архив'
			files = os.listdir(f"{path}/{temp[n][1]['name']}")
			cert_list = []
			container_list = []
			zip_list = []
			print('_'*30)
			print(f"Подпишите «{temp[n][1]['name']}» [{n+1}/{len(temp)}]")
			time.sleep(1)
			for i in files:
				if i.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
					container_list.append(i)
				if i.endswith('.cer'):
					cert_list.append(i)
					cert_current = 1
				if i.endswith('.zip'):
					zip_list.append(i)
					zip_current = 1
			if len(container_list) == 0:
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствует %s!" % container_current)
					time.sleep(1)
					msg = input(f"\nПовторно проверить наличие контейнера - (нажать Enter) | Пропустить - (ввести 2)\n")
					if msg == '':
						animation(2, 'Поиск контейнера')
						continue
					if msg == '2':
						print()
						break
			if len(container_list) > 0:
				if eds[temp[n][1]['name']]['EDS'] == '':
					eds[temp[n][1]['name']]['EDS'] = container_list[-1]
					try:
						os.mkdir('D:\\' + container_list[-1])
						copytree(path + temp[n][1]['name'] + '/' + container_list[-1], 'D:\\' + container_list[-1])
						try:
							print(f"{temp[n][1]['name']}: {container_list[-1]} --- {eds[temp[n][1]['name']]['password']}")
						except KeyError:
							pass
						dirs_on_D.append((temp[n][1]['name'], container_list[-1]))
						#TODO!
						
						
						'''except PermissionError:
							not_container = []
							try:
								with open('EDS_not_container.txt', 'r') as file_nc:
									for i in file_nc:
										not_container.append(i.strip())
							except:
								with open('EDS_not_container.txt', 'w') as f:
									pass
							with open('EDS_not_container.txt', 'a') as f:
								if value['name'] not in not_container:
									f.write(f"{value['name']}\n")
							print(f"{value['name']}: Папка НЕ СКОПИРОВАНА на диск D:\. Отсутствует контейнер")
							pass'''
					except FileExistsError:
						if value['EDS'] not in folders_sign_to_del:
							shutil.rmtree('D:\\' + container_list[-1])
							os.mkdir('D:\\' + container_list[-1])
							copytree(path + temp[n][1]['name'] + '/' + container_list[-1], 'D:\\' + container_list[-1])
						try:
							print(f"{temp[n][1]['name']}: {container_list[-1]} --- {eds[temp[n][1]['name']]['password']}")
						except KeyError:
							pass
							dirs_on_D.append((temp[n][1]['name'], container_list[-1]))
						else:
							dirs_on_D.append((temp[n][1]['name'], container_list[-1]))
			
			if temp[n][1]['password'] != '':
				pass_current = 1
			if cert_current != 1 and pass_current != 1:
				if zip_current != 1:
					delFromClipboard()
					time.sleep(1.5)
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствуют %s и %s!" % (cert_current, pass_current))
					time.sleep(1)
					msg = input(f"\nПовторно проверить сертификат - (нажать Enter) | Пропустить - (ввести 2)\n")
					if msg == '':
						continue
					if msg == '2':
						print()
						break
				elif zip_current == 1:
					delFromClipboard()
					time.sleep(1.5)
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствуют %s и %s!"
							f"\nНо найден zip-архив с сертификатом!" % (cert_current, pass_current))
					time.sleep(1)
					msg = input(f"\nПовторно проверить сертификат - (нажать Enter) | "
								f"Распаковать zip-архив - (ввести 2) | Пропустить - (ввести 3)\n")
					if msg == '':
						continue
					if msg =='2':
						with zipfile.ZipFile(f"{path}/{temp[n][1]['name']}/{zip_list[-1]}", 'r') as zip_file:
							zip_file.extractall(f"{path}/{temp[n][1]['name']}")
						continue
					
					if msg == '3':
						print()
						break
					
					
			elif cert_current != 1:
				if zip_current != 1:
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствует %s!" % cert_current)
					time.sleep(1)
					msg = input(f"\nПовторно проверить сертификат - (нажать Enter) | Пропустить - (ввести 2)\n")
					delFromClipboard()
					if msg == '':
						continue
					if msg == '2':
						print()
						break
				elif zip_current == 1:
					delFromClipboard()
					time.sleep(1.5)
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствует %s!"
							f"\nНо найден zip-архив с сертификатом!" % cert_current)
					time.sleep(1)
					msg = input(f"\nПовторно проверить сертификат - (нажать Enter) | "
								f"Распаковать zip-архив - (ввести 2) | Пропустить - (ввести 3)\n")
					if msg == '':
						continue
					if msg =='2':
						with zipfile.ZipFile(f"{path}/{temp[n][1]['name']}/{zip_list[-1]}", 'r') as zip_file:
							zip_file.extractall(f"{path}/{temp[n][1]['name']}")
						continue
					
					if msg == '3':
						print()
						break

			elif cert_current == 1:
				remained_current = period(eds[temp[n][1]['name']]['name'], cert_list[-1])['remained']
				if pass_current != 1:
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» отсутствует %s!" % pass_current)
					time.sleep(1.5)
				
				print(f"Период действия  c: {period(eds[temp[n][1]['name']]['name'], cert_list[-1])['begin']}"
						f"\n\t\tдо: {period(eds[temp[n][1]['name']]['name'], cert_list[-1])['end']}"
						f"\n\t      ост.: {days_remained(period(eds[temp[n][1]['name']]['name'], cert_list[-1])['remained'])}")
				time.sleep(1)
				if int(remained_current) < 0:
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» истёк срок действия сертификата!")
					time.sleep(1)
					msg = input(f"\nПовторно проверить сертификат - (нажать Enter) | Пропустить - (ввести 2)\n")
					delFromClipboard()
					if msg == '':
						continue
					if msg == '2':
						print()
						break
				'''print(f"Действует с: {period(eds[temp[n][1]['name']]['name'], cert_list[-1])['begin']}; "
					f"Действует до: {period(eds[temp[n][1]['name']]['name'], cert_list[-1])['end']}; "
					f"Осталось дней: {period(eds[temp[n][1]['name']]['name'], cert_list[-1])['remained']}")'''
				if int(remained_current) in range (0,46):
					print(f"ВНИМАНИЕ!!! У сотрудника «{temp[n][1]['name']}» истекает срок действия сертификата!")
					time.sleep(1.5)
				addToClipBoard(f"{path}{temp[n][1]['name']}/{cert_list[-1]}")
				os.startfile (r"D:\EDS_CP.lnk")
				time.sleep(5)
				if pass_current == 1:
					addToClipBoard(temp[n][1]['password'])
				elif pass_current != 1:
					pass
					#delFromClipboard()
				if len(temp) != n + 1:
					msg = input('\nНажмите Enter, чтобы перейти к следующему сотруднику\n')
					if msg == '':
						pass
			break
os.system('pause')
if len(dirs_on_D) > 0:
	print()
	animation(2, 'Отправка')
else:
	print('\nОтсутсвуют ЭЦП для отправки!')
	time.sleep(0.5)

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None  # disable host key checking.

# Подключение к SFTP-серверу
with pysftp.Connection(host='103.58.1.13',
                       port=22,
                       username=username,
                       password=password,
                       cnopts=cnopts,
					   log=path + backup_dir + "/sftp.log"
                       ) as sftp:

    # Сменить директорию на SFTP-сервере
	with sftp.cd('/home/' + username):
		with open ('EDS_log.txt', 'a') as file:
			for i,j in dirs_on_D:
				files = os.listdir('D:\\' + j)
				while True:
					cert_current = 'сертификат'
					pass_current = 'пароль от подписи'
					container_current = 'контейнер'
					zip_current = 'zip-архив'
					if Path(os.getcwd() + j +  '/header.key').stat().st_size < 1500:
						print(f'\n{i}: {j} НЕ ПЕРЕДАН, т.к. сертификат не был установлен через CryptoPro!')
						msg = input(f'Подписать и отправить - (нажать Enter) | Повторить отправку - (ввести 2) | Пропустить отправку - (ввести 3)\n')
						
						if msg == '':
							files_temp = os.listdir(f"{path}/{i}")
							cert_list = []
							zip_list = []
							for f in files_temp:
								if f.endswith('.cer'):
									cert_list.append(f)
									cert_current = 1
								if f.endswith('.zip'):
									zip_list.append(f)
									zip_current = 1
							if eds[i]['password'] != '':
								pass_current = 1
							if cert_current != 1 and pass_current != 1:
								if zip_current != 1:
									delFromClipboard()
									time.sleep(1.5)
									print(f"ВНИМАНИЕ!!! У сотрудника «{i}» отсутствуют %s и %s!" % (cert_current, pass_current))
									time.sleep(2)
									continue
								elif zip_current == 1:
									delFromClipboard()
									time.sleep(1.5)
									with zipfile.ZipFile(f"{path}/{i}/{zip_list[-1]}", 'r') as zip_file:
											zip_file.extractall(f"{path}/{i}")
									for h in os.listdir(f"{path}/{i}"):
										if h.endswith('.cer'):
											cert_list.append(h)
									remained_current =  period(i, cert_list[-1])['remained']		
									
									if int(remained_current) < 0:
										print(f"ВНИМАНИЕ!!! У сотрудника «{i}» истёк срок действия сертификата!")
										time.sleep(2)
										delFromClipboard()
										break
									if int(remained_current) in range (0,46):
										print(f"ВНИМАНИЕ!!! У сотрудника «{i}» истекает срок действия сертификата!")
										time.sleep(2)
						
									print(f"ВНИМАНИЕ!!! У сотрудника «{i}» отсутствуют %s и %s!"
										f"\nНо найден и РАСПАКОВАН zip-архив с сертификатом! Повторно выберите 'Подписать и отправить'" % (cert_current, pass_current))
									time.sleep(2)
									print(f"\n«{i}»\nПериод действия  c: {period(i, cert_list[-1])['begin']}"
										f"\n\t\tдо: {period(i, cert_list[-1])['end']}"
										f"\n\t      ост.: {days_remained(period(i, cert_list[-1])['remained'])}")
									time.sleep(2)	

							elif cert_current != 1:
								if zip_current != 1:
									delFromClipboard()
									time.sleep(1.5)
									print(f"ВНИМАНИЕ!!! У сотрудника «{i}» отсутствует %s!" % cert_current)
									time.sleep(2)
									continue
								elif zip_current == 1:
									delFromClipboard()
									time.sleep(1.5)
									with zipfile.ZipFile(f"{path}/{i}/{zip_list[-1]}", 'r') as zip_file:
											zip_file.extractall(f"{path}/{i}")
									for h in os.listdir(f"{path}/{i}"):
										if h.endswith('.cer'):
											cert_list.append(h)
									remained_current =  period(i, cert_list[-1])['remained']		
									if int(remained_current) < 0:
										print(f"ВНИМАНИЕ!!! У сотрудника «{i}» истёк срок действия сертификата!")
										time.sleep(2)
										delFromClipboard()
										break
									if int(remained_current) in range (0,46):
										print(f"ВНИМАНИЕ!!! У сотрудника «{i}» истекает срок действия сертификата!")
										time.sleep(2)
									print(f"ВНИМАНИЕ!!! У сотрудника «{i}» отсутствует %s!"
										f"\nНо найден и РАСПАКОВАН zip-архив с сертификатом! Повторно выберите 'Подписать и отправить'" % cert_current)
									time.sleep(2)
								
									print(f"\n«{i}»\nПериод действия  c: {period(i, cert_list[-1])['begin']}"
										f"\n\t\tдо: {period(i, cert_list[-1])['end']}"
										f"\n\t      ост.: {days_remained(period(i, cert_list[-1])['remained'])}")
									time.sleep(2)
									
							elif cert_current == 1:
								if pass_current !=1:
									print(f"ВНИМАНИЕ!!! У сотрудника «{i}» отсутствует %s!" % pass_current)
									time.sleep(1.5)
								addToClipBoard(f"{path}/{i}/{cert_list[-1]}")
								os.startfile (r"D:\EDS_CP.lnk")
								time.sleep(5)
								if pass_current == 1:
									addToClipBoard(temp[n][1]['password'])
								elif pass_current != 1:
									pass
									#delFromClipboard()
								send_msg = input('Нажмите Enter, чтобы отправить\n')
								continue
						elif msg =='2':
							continue
						elif msg == '3':
							file.write(f"{i}: {j} НЕ ПЕРЕДАН, т.к. сертификат не был установлен через CryptoPro!\n")
							break
						else:
							continue
					else:
						break
				if Path(os.getcwd() + j +  '/header.key').stat().st_size < 1500:
					print(f'{i}: {j} ПРОПУЩЕН')
				elif Path(os.getcwd() + j +  '/header.key').stat().st_size >= 1500:
					time.sleep(2)
					sftp.mkdir(j)
					time.sleep(2)
					for k in files:
						sftp.put('D:\\' + j + '/' + k, '/home/' + username + '/' + j + '/' + k)
					print(f'{i}: {j} передан')
					file.write(f"{i}: {j} передан\n")
					try:
						os.mkdir(f"{path}{backup_dir}/{j} {i}")
						copytree('D:\\' + j, f"{path}{backup_dir}/{j} {i}")
					except FileExistsError:
						shutil.rmtree(f"{path}{backup_dir}/{j} {i}")
						time.sleep(0.5)
						os.mkdir(f"{path}{backup_dir}/{j} {i}")
						copytree('D:\\' + j, f"{path}{backup_dir}/{j} {i}")
					except PermissionError:
						time.sleep(1.5)
						try:
							shutil.rmtree(f"{path}{backup_dir}/{j} {i}")
						except:
							os.mkdir(f"{path}{backup_dir}/{j} {i}")
							copytree('D:\\' + j, f"{path}{backup_dir}/{j} {i}")

os.system('pause')