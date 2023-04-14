import pandas as pd
import glob
import warnings
import datetime
import win32com.client
warnings.simplefilter("ignore")

def macro():
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(Filename=r"C:\Users\2\Desktop\Отчётность Яндекс доставка\Отчеты\Отчёт ЯД - апрель.xlsm", UpdateLinks=True)
    excel.Application.Run("refresh_report")
    workbook.Close()
    excel.Quit()
    del workbook
    del excel

def yd_report():
    files = [item for item in glob.glob(r'C:\Users\2\Desktop\Отчётность Яндекс доставка\ЯФ выгрузка\*{}'.format('.xlsx'))]
    count=0
    yd = pd.DataFrame()
    for file in files:
        file=pd.read_excel(file)
        yd = pd.concat([yd, file])
        count+=1
    yd['Время создания']=pd.to_datetime(yd['Время создания'], format = "%Y-%m-%d %H:%M:%S").dt.date
    yd = yd.drop(columns= ['Сотрудник'], axis = 1)
    yd = yd.rename(columns= {'Время создания' : 'Дата', '<font color=black choice_id=1>Сотрудник' : 'Сотрудник', '<font color=207567>\n\n\nЗдравствуйте! Звоню, потому что проводим оптимизацию сервиса яндекс карты, и у вас не указана информация о доставке, соедините с логистом или директором для актуализации! \n\nЗдравствуйте! Сейчас проводим оптимизацию сервиса яндекс карты, и у вас не указана информация о доставке, с кем могу обсудить данный вопрос?\n\nЗдравствуйте! Сейчас проводим оптимизацию сервиса яндекс карты, у вас не указана информация о доставке, соедините с директором для утонения информации.\n\nЗдравствуйте! Звоню по сервису Я. карты, оптимизируем сервис по интеграции с доставкой, соедините с логистом или директором!' : 'Дозвон?'})


    files_sk = [item for item in glob.glob(r'C:\Users\2\Desktop\Отчётность Яндекс доставка\Скорозвон выгрузки\*{}'.format('.xlsx'))]
    count_sk=0
    sk = pd.DataFrame()
    for file in files_sk:
        file=pd.read_excel(file)
        sk = pd.concat([sk, file] )
        count_sk+=1
    sk['Дата']=pd.to_datetime(sk['Дата'], format = '%d.%m.%Y').dt.date
    sk = sk[sk['Сотрудник'] !='(без ответственного)']


    # Гугл док

    google = pd.read_csv(f'https://docs.google.com/spreadsheets/d/1A9xoKgiVSRBV2qEJWIXaeOzWauBOT7yLsWQOpNRGP4g/gviz/tq?tqx=out:csv&gid=285132865')
    google = google.rename(columns = {'ФИО' : 'Оператор'})
    google['Оператор'] = google['Оператор'].map(str.strip)
    google = google.drop_duplicates(subset = ['Оператор'])
    google = google.loc[:, ['Оператор', 'РГ', 'Стаж', 'Офис/удаленка']]
    google = google.rename(columns= {'Оператор' : 'Сотрудник'})



    # Скорозвон
    sk_all_calls = sk.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID лида' : 'count'}).rename(columns= {'ID лида' : 'Всего звонков скр'})
    sk_unique = sk.drop_duplicates(subset = ['Дата', 'Сотрудник', 'Телефон, на который звонили'])
    sk_unique = sk_unique.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID лида' : 'count'}).rename(columns= {'ID лида' : 'Уникал скр'})
    sk_success = sk.drop_duplicates(subset = ['Дата', 'Сотрудник', 'Телефон, на который звонили'])
    sk_success = sk_success[(sk_success['Результат'] == 'Работа без договора') | (sk_success['Результат'] == 'Заключение договора')]
    sk_success = sk_success.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID лида' : 'count'}).rename(columns= {'ID лида' : 'Успех скр'})
    sk = sk_all_calls.merge(sk_unique, how = 'left', on = ['Дата', 'Сотрудник']).merge(sk_success, how = 'left', on = ['Дата', 'Сотрудник'])


    # Яндекс доставка
    yd_all_calls = yd.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Всего звонков ЯФ'})
    yd_unique = yd.drop_duplicates(subset = ['Дата', 'Сотрудник', 'Телефон'])
    yd_unique = yd_unique.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Уникал ЯФ'})
    # Дозвон
    yd_dozvon = yd[((yd['Статус звонка'] == 'Звонок состоялся') | (yd['Статус звонка'] == 'Бросили трубку')) & ((yd['Дозвон?'] == 'Со мной / Дозвон до ЛПР') | (yd['Дозвон?'] == 'Не удалось получить контакт ЛПР / Отказ от сервиса') | (yd['Дозвон?'] == 'Попросили перезвонить')) & (yd['Причина отказа от сервиса'] != 'Возможности сервиса не соответствуют запросу клиента')]
    yd_dozvon = yd_dozvon.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Дозвон ЯФ'})
    # Дозвон без нецелевых
    yd_dozvon_without_aim = yd[(yd['Дозвон?'] == 'Со мной / Дозвон до ЛПР') | (yd['Дозвон?'] == 'Не удалось получить контакт ЛПР / Отказ от сервиса') | (yd['Дозвон?'] == 'Попросили перезвонить')]
    yd_dozvon_without_aim = yd_dozvon_without_aim.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Дозвон без нецелевых ЯФ'})
    # Количество нецелевых клиентов
    yd_aim = yd[yd['Дозвон?'] == 'Нецелевой клиент']
    yd_aim = yd_aim.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Кол-во нецелевых ЯФ'})
    yd_success = yd.drop_duplicates(subset = ['Дата', 'Сотрудник', 'Телефон'])
    yd_success = yd_success[(yd_success['<font color=207567>Подскажите, какой вариант доставок больше подходит для вашего бизнеса?'].notna()) & (yd_success['<font color=207567>Подскажите, какой вариант доставок больше подходит для вашего бизнеса?'] != 'Вообще не готов делать доставку')]
    yd_success = yd_success.groupby(['Дата', 'Сотрудник'], as_index= False).agg({'ID' : 'count'}).rename(columns= {'ID' : 'Успех ЯФ'})
    yd = yd_all_calls.merge(yd_unique, how = 'left', on = ['Дата', 'Сотрудник']).merge(yd_dozvon, how = 'left', on = ['Дата', 'Сотрудник']).merge(yd_dozvon_without_aim, how = 'left', on = ['Дата', 'Сотрудник']).merge(yd_aim, how = 'left', on = ['Дата', 'Сотрудник']).merge(yd_success, how = 'left', on = ['Дата', 'Сотрудник'])

    # Скорозвон + ЯФ
    sk_yd = sk.merge(yd, how = 'left', on = ['Дата', 'Сотрудник'])

    # Отчет по дням по операторам
    day_op = sk_yd.merge(google, how = 'left', on = ['Сотрудник'])
    total_day = day_op
    total_day = total_day.fillna({'РГ' : 'Нет данных', 'Стаж' : 'Нет данных'})
    total = total_day

    # Накопительный по РГ
    total = total_day.groupby(['РГ'], as_index= False).agg({'Всего звонков скр' : 'sum', 'Уникал скр' : 'sum', 'Успех скр' : 'sum', 'Всего звонков ЯФ' : 'sum', 'Уникал ЯФ' : 'sum', 'Дозвон ЯФ' : 'sum', 'Дозвон без нецелевых ЯФ' : 'sum', 'Кол-во нецелевых ЯФ' : 'sum', 'Успех ЯФ' : 'sum', 'Сотрудник' : lambda x: len(pd.unique(x))}).rename(columns= {'Сотрудник' : 'fte'})
    total['Конверсия уникал скр'] = total['Успех скр']/total['Уникал скр']
    total['Конверсия уникал ЯФ'] = total['Успех ЯФ']/total['Уникал ЯФ']
    total['Конверсия дозвона ЯФ'] = total['Успех ЯФ']/total['Дозвон ЯФ']
    total['Конверсия дозвона без нецелевых ЯФ'] = total['Успех ЯФ']/total['Дозвон без нецелевых ЯФ']


    # Конверсии для отчет по дням по операторам
    day_op['Конверсия уникал скр'] = day_op['Успех скр']/day_op['Уникал скр']
    day_op['Конверсия уникал ЯФ'] = day_op['Успех ЯФ']/day_op['Уникал ЯФ']
    day_op['Конверсия дозвона ЯФ'] = day_op['Успех ЯФ']/day_op['Дозвон ЯФ']
    day_op['Конверсия дозвона без нецелевых ЯФ'] = day_op['Успех ЯФ']/day_op['Дозвон без нецелевых ЯФ']


    # Отчет по РГ по дням с конверсиями
    total_day = total_day.groupby(['Дата', 'РГ'], as_index= False).agg({'Всего звонков скр' : 'sum', 'Уникал скр' : 'sum', 'Успех скр' : 'sum', 'Всего звонков ЯФ' : 'sum', 'Уникал ЯФ' : 'sum', 'Дозвон ЯФ' : 'sum', 'Дозвон без нецелевых ЯФ' : 'sum', 'Кол-во нецелевых ЯФ' : 'sum', 'Успех ЯФ' : 'sum', 'Сотрудник' : 'count'}).rename(columns= {'Сотрудник' : 'fte'})
    total_day['Конверсия уникал скр'] = total_day['Успех скр']/total_day['Уникал скр']
    total_day['Конверсия уникал ЯФ'] = total_day['Успех ЯФ']/total_day['Уникал ЯФ']
    total_day['Конверсия дозвона ЯФ'] = total_day['Успех ЯФ']/total_day['Дозвон ЯФ']
    total_day['Конверсия дозвона без нецелевых ЯФ'] = total_day['Успех ЯФ']/total_day['Дозвон без нецелевых ЯФ']


    # Накопительный по операторам с конверсиями
    all_op = sk_yd.groupby('Сотрудник', as_index= False).agg({'Всего звонков скр' : 'sum', 'Уникал скр' : 'sum', 'Успех скр' : 'sum', 'Всего звонков ЯФ' : 'sum', 'Уникал ЯФ' : 'sum', 'Дозвон ЯФ' : 'sum', 'Дозвон без нецелевых ЯФ' : 'sum', 'Кол-во нецелевых ЯФ' : 'sum', 'Успех ЯФ' : 'sum'}).merge(google, how = 'left', on = ['Сотрудник'])
    all_op['Конверсия уникал скр'] = all_op['Успех скр']/all_op['Уникал скр']
    all_op['Конверсия уникал ЯФ'] = all_op['Успех ЯФ']/all_op['Уникал ЯФ']
    all_op['Конверсия дозвона ЯФ'] = all_op['Успех ЯФ']/all_op['Дозвон ЯФ']
    all_op['Конверсия дозвона без нецелевых ЯФ'] = all_op['Успех ЯФ']/all_op['Дозвон без нецелевых ЯФ']


    # СНГ по дням
    cis_day = yd.merge(google, how = 'left', on = ['Сотрудник'])
    # СНГ накопительный с конверсиями
    all_cis = cis_day.groupby(['Сотрудник'], as_index= False).agg({'Всего звонков ЯФ' : 'sum', 'Уникал ЯФ' : 'sum', 'Дозвон ЯФ' : 'sum', 'Дозвон без нецелевых ЯФ' : 'sum', 'Кол-во нецелевых ЯФ' : 'sum', 'Успех ЯФ' : 'sum'}).merge(google, how = 'left', on = ['Сотрудник'])
    all_cis = all_cis[(all_cis['РГ'] == 'Армения') | (all_cis['РГ'] == 'Беларусь') | (all_cis['РГ'] == 'Казахстан')]
    all_cis['Конверсия всего звонков ЯФ'] = all_cis['Успех ЯФ']/all_cis['Всего звонков ЯФ']
    all_cis['Конверсия уникал ЯФ'] = all_cis['Успех ЯФ']/all_cis['Уникал ЯФ']
    all_cis['Конверсия дозвона ЯФ'] = all_cis['Успех ЯФ']/all_cis['Дозвон ЯФ']
    all_cis['Конверсия дозвона без нецелевых ЯФ'] = all_cis['Успех ЯФ']/all_cis['Дозвон без нецелевых ЯФ']


    # Конверсии для отчета по дням СНГ
    cis_day = cis_day[(cis_day['РГ'] == 'Армения') | (cis_day['РГ'] == 'Беларусь') | (cis_day['РГ'] == 'Казахстан')]
    cis_day['Конверсия всего звонков ЯФ'] = cis_day['Успех ЯФ']/cis_day['Всего звонков ЯФ']
    cis_day['Конверсия уникал ЯФ'] = cis_day['Успех ЯФ']/cis_day['Уникал ЯФ']
    cis_day['Конверсия дозвона ЯФ'] = cis_day['Успех ЯФ']/cis_day['Дозвон ЯФ']
    cis_day['Конверсия дозвона без нецелевых ЯФ'] = cis_day['Успех ЯФ']/cis_day['Дозвон без нецелевых ЯФ']

    with pd.ExcelWriter(r"C:\Users\2\Desktop\Отчётность Яндекс доставка\Отчеты пандас\Отчет ЯД.xlsx") as writer:
        day_op.to_excel(writer, sheet_name='По дням операторы', index = False)
        all_op.to_excel(writer, sheet_name='Накоп операторы', index = False)
        total_day.to_excel(writer, sheet_name='По дням РГ', index = False)
        total.to_excel(writer, sheet_name='Накоп РГ', index = False)
        cis_day.to_excel(writer, sheet_name='По дням АБК', index = False)
        all_cis.to_excel(writer, sheet_name='Накоп АБК', index = False)

    del yd, sk, yd_all_calls, yd_unique, yd_dozvon_without_aim, yd_dozvon, sk_all_calls, sk_unique, day_op, all_op, sk_success, yd_aim, yd_success, sk_yd, cis_day, all_cis, total_day, total
    print("Запись в файл: все ровно ЯД обычный")
    macro()
    print("Макрос: все ровно ЯД обычный")

if __name__ == "__main__":
    yd_report()