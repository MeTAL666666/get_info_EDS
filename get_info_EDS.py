import os, sys, re, csv, time, locale
from datetime import datetime



locale.setlocale(
    category=locale.LC_ALL,
    locale="Russian"
)

path = "//Administrator/общая/ЭЦП/"

try:
	os.path.isdir(path)
except FileNotFoundError:
	path = 'D:\общая\ЭЦП/'

'''Функция для копирования в буфер обмена!'''
'''def addToClipBoard(text):
    command = 'echo ' + text.strip() + '| clip'
    os.system(command)

addToClipBoard('test')'''

def write_to_csv(data):
	if os.path.exists(path + '!EXPORT.csv'):
		os.remove(path + '!EXPORT.csv')
	with open (path + '!EXPORT.csv', 'w', encoding='utf-8') as file:
		writer = csv.writer	(file, delimiter='\t')
		data['headers']['dt_create_file'] = datetime.strftime(datetime.now(),'%d.%m.%y %a %H:%M')
		for i in data:
			writer.writerow((data[i]['name'],
							data[i]['name_key'],
							data[i]['password'],
							data[i]['eds_zip'],
							data[i]['eds_cer'],
							data[i]['eds_dir'],
							data[i]['begin'],
							data[i]['end'],
							data[i]['remained'],
							data[i]['dt_create_file'],
							data[i]['dt_now'],
							data[i]['refresh'],
							data[i]['get_not_sign'],
							data[i]['repeated_container'],
							data[i]['remained_true']))

	print('\n' + '-'*20 + '\nСписок создан!')

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

	begin_str = begin.strftime('%d.%m.%y')
	end_str = end.strftime('%d.%m.%y')
	remained_str = str(end - datetime.now()).split(' ')[0]
	return {'begin': begin_str, 'end': end_str, 'remained': remained_str}

main_dirs = [i for i in os.listdir(path) if os.path.isdir(path + i) and '!' not in i and i not in [j for j in ('ХЗ','смена фамилии', 'Заявы', 'Заявления', 'backup_sent_EDS')]]

eds = {'headers':{'name': '=ГИПЕРССЫЛКА("D:\\EDS.exe";"ФИО(+ run EDS.exe)")',
					'name_key': 'Название подписи',
					'password': 'Пароль',
					'eds_zip': 'Сертификат ЭЦП',
					'eds_cer':'Распакован',
					'eds_dir': 'Контейнер',
					'begin': 'Действует С',
					'end': 'Действует ДО',
					'remained': 'Осталось (дней)',
					'dt_create_file': '',
					'dt_now': '=СЕГОДНЯ()',
					'refresh': '=ГИПЕРССЫЛКА("get_info_EDS.exe";"Пересобрать список")',
					'get_not_sign': '=ГИПЕРССЫЛКА("sftp_not_sign.exe";"Найти неподписанные контейнеры на сервере")',
					'repeated_container': '=СЧЁТЕСЛИ($N$2:$N$1734;"Дубликат")',
					'remained_true': '=СЧЁТЕСЛИ($O$2:$O$1734;ИСТИНА)'}}


for n, i in enumerate(main_dirs):
	f = n + 2
	list_files = os.listdir(path + i)
	link = ('=ГИПЕРССЫЛКА("%s\\";"%s")' %(i, i))
	eds[i] = {'name':link,
				'eds_dir': '',
				'eds_zip':'',
				'eds_cer':'' ,
				'name_key': '',
				'password':'',
				'begin':'',
				'end':'',
				'remained':'',
				'dt_create_file':'',
				'dt_now':'',
				'refresh':'',
				'get_not_sign':'',
				'repeated_container': '=ЕСЛИ(СЧЁТЕСЛИ($F$2:$F$1734; F%s)>1;"Дубликат";"")' % f ,
				'remained_true': '=ЕСЛИ(I%s<>"";ЕСЛИ(I%s<=45;ИСТИНА;"");"")' % (f,f) }

	print(f"[{n+1} / {len(main_dirs)}] {i}...OK! " + ' ' * 40,end='\r')
	
	for j in list_files:
		if j.endswith(tuple(['.' +  str(j)  + '0' + str(i) for j in range(10) for i in range(10)])):
			eds[i]['eds_dir'] = j
		if j.endswith(tuple(['.zip', '.cer'])):
			eds[i]['eds_zip'] = f'=ГИПЕРССЫЛКА("D:\\EDS_CP.lnk";"//Administrator/общая/ЭЦП/{i}/{j.replace(".zip",".cer")}")'
		if j.lower().endswith('.cer'):
			eds[i]['eds_cer'] = 'unzip'
			eds[i]['begin'] = period(i, j)['begin']
			eds[i]['end'] = period(i, j)['end']
			eds[i]['remained'] = f'=ОКРУГЛ(ДАТАЗНАЧ("{eds[i]["end"]}") - $K$1;0)'
		if j.endswith('.txt'):
			try:
				with open (path + i + '/' + j) as f:
					password = re.findall(r'\n\d{5,6}',f.read())[-1].strip()
					eds[i]['password'] = password
			except:
				eds[i]['password'] = ''
	try:
		with open (path + i + '/' + eds[i]['eds_dir'] + '/name.key', encoding='cp1251') as f:
				eds[i]['name_key'] = f.read()[4:]
	except:
		eds[i]['name_key'] = ''

count_run = 0
while True:
	try:
		write_to_csv(eds)
		break
	except PermissionError:
		msg = input(f'\nФайл !EXPORT.csv открыт. Закройте файл и нажмите Enter\n')
		if msg =='':
			continue
		else:
			continue
os.system('pause')