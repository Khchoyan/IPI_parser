import pandas as pd
import re
import requests
from bs4 import BeautifulSoup as bs
import os
import time
import datetime
from calendar import monthrange


def reformate_date(date, year):
    date = re.sub('[0-9]', '', date).strip().lower()
    flag = True if ((year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)) else False
    if date == 'январь':
        date = '31 january'
    elif date == 'февраль' or date == 'январь-февраль' and flag:
        date = '29 february'
    elif date == 'февраль' or date == 'январь-февраль':
        date = '28 february'
    elif date == 'март' or date == 'январь-март':
        date = '31 march'
    elif date == 'апрель' or date == 'январь-апрель':
        date = '30 April'
    elif date == 'май' or date == 'январь-май':
        date = '31 may'
    elif date == 'июнь' or date == 'январь-июнь':
        date = '30 june'
    elif date == 'июль' or date == 'январь-июль':
        date = '31 july'
    elif date == 'август' or date == 'январь-август':
        date = '31 august'
    elif date == 'сентябрь' or date == 'январь-сентябрь':
        date = '30 september'
    elif date == 'октябрь' or date == 'январь-октябрь':
        date = '31 october'
    elif date == 'ноябрь' or date == 'январь-ноябрь':
        date = '30 november'
    elif date == 'декабрь' or date == 'январь-декабрь':
        date = '31 december'
    return date


def pars_year_by_months():
    """
    Функция для получения ссылок на документы по месяцам.
    Для ВВП реализовано возвращение названия последнего доступного месяца в конкретном году
    и ссылки на соответствующий раздел.
    """
    header = {
        'user-agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:86.0) Gecko/20100101 Firefox/86.0'
    }
    time.sleep(15)
    url = f'https://rosstat.gov.ru/enterprise_industrial#'
    response = requests.get(url, headers=header)
    soup = bs(response.content, "html.parser")
    for i in soup.find_all('a', {'class': "btn btn-icon btn-white btn-br btn-sm"}):
        if '/storage/mediabank/ind_sub_2018.xlsx' in str(i.get('href')):
            link_to_download = f'https://rosstat.gov.ru' + str(i.get('href'))
            print(link_to_download)
            time.sleep(15)
            dok_name_to_download = 'file.xlsx'
            folder = os.getcwd()
            response = requests.get(link_to_download, headers=header)
            folder = os.path.join(folder, 'word_data', dok_name_to_download)
            if response.status_code == 200:
                with open(folder, 'wb') as f:
                    f.write(response.content)
                print(f'Document was downloaded.')
            else:
                print('FAILED:', link_to_download)
            return 'word_data/' + dok_name_to_download


def create_dict(df, year):
    dct = {}
    date = list(df.loc[3])
    value = list(df.loc[4])
    count = 10**6
    for i in range(len(df.iloc[0])):
        if ('январь' in date[i].lower() and i != 0 and 'январь-' not in date[i].lower()) or i >= count:
            dct[reformate_date(date[i], year) + str(year)] = value[i]
            count = i
        else:
            dct[pd.to_datetime(reformate_date(date[i], year - 1) + str(year - 1))] = value[i]
    return dct


def create_new_date(last_date_in_file_year, last_date_in_file_month):
    now = datetime.datetime.now()
    lst_date = []
    _, last_day = monthrange(now.year, now.month)
    last_date = datetime.datetime.strptime(f"{now.year}-{now.month}-{last_day}", "%Y-%m-%d").date()
    for i in range((last_date.year - last_date_in_file_year) * 12 + last_date.month - last_date_in_file_month - 1):
        if last_date.month - 1 != 0:
            _, last_day = monthrange(last_date.year, last_date.month - 1)
            last_date = datetime.datetime.strptime(f"{last_date.year}-{last_date.month - 1}-{last_day}",
                                                   "%Y-%m-%d").date()
        else:
            _, last_day = monthrange(last_date.year - 1, 12)
            last_date = datetime.datetime.strptime(f"{last_date.year - 1}-{12}-{last_day}", "%Y-%m-%d").date()
        lst_date.append(last_date)
    return sorted(lst_date)


def append_date_rez_file_Y(xlsx_path='rez_file_Y_v2.xlsx'):
    """
        Функция осуществляет дабавление месяцев, если их нет в файле.
    """
    data_xlsx = pd.read_excel(xlsx_path)
    year = pd.to_datetime(pd.read_excel('rez_file_Y_v2.xlsx')['Целевой показатель'].iloc[-1]).year
    month = pd.to_datetime(pd.read_excel('rez_file_Y_v2.xlsx')['Целевой показатель'].iloc[-1]).month
    date_lst = create_new_date(year, month)
    for date in date_lst:
        new_string = {'Целевой показатель': [date]}
        new_string.update({c: [None] for c in data_xlsx.columns[1:]})
        new_string = pd.DataFrame(new_string)
        if not data_xlsx.empty and not new_string.empty:
            data_xlsx = pd.concat([data_xlsx, new_string])
    data_xlsx.to_excel(xlsx_path, index=False)


def update_rez_file_y(data, column_name, xlsx_path='rez_file_Y_v2.xlsx'):
    """
        Функция осуществляет обновление файла со всеми данными rez_file_Y_v2.xlsx
    """
    data_xlsx = pd.read_excel(xlsx_path)
    if list(data.keys())[-1] not in list(data_xlsx['Целевой показатель']):
        append_date_rez_file_Y()
        data_xlsx = pd.read_excel(xlsx_path)

    for j in data:
        data_xlsx.loc[data_xlsx['Целевой показатель'] == j, column_name] = data[j]

    data_xlsx.to_excel(xlsx_path, index=False)
    print(f'rez_file_Y_v2.xlsx was apdated')


def main():
    year = datetime.datetime.now().year
    path = pars_year_by_months()
    # path = 'word_data/file.xlsx'
    df_1 = create_dict(pd.read_excel(path, sheet_name='1').loc[[3, 4]].iloc[:, -12:], year)
    df_2 = create_dict(pd.read_excel(path, sheet_name='2').loc[[3, 4]].iloc[:, -12:], year)
    df_3 = create_dict(pd.read_excel(path, sheet_name='3').loc[[3, 4]].iloc[:, -12:], year)

    update_rez_file_y(df_1, 'ИПП в % к соответствующему месяцу предыдущего года')
    update_rez_file_y(df_2, 'ИПП в % к соответствующему периоду предыдущего года')
    update_rez_file_y(df_3, 'ИПП в % к предыдущему месяцу')


if __name__ == '__main__':
    main()
