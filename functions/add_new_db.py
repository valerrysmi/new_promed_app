import pandas as pd
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import re
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell

from general_vars import general_vars

def create_sheet_names(today):
    sheet_names = []
    year_now = today.year

    for x in range(year_now - 18, year_now + 1):
        sheet_names.append(str(x))

    return sheet_names


def year_month_now(today):
    return str(today.year) + '/' + ('0' + str(today.month))[-2:]


def read_data(file_path, sheet_names, min_bd, act_year_month):
    print('read_data')
    month_names = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
    data = pd.read_excel(file_path, sheet_name=sheet_names)
    print(data.keys())
    df = pd.DataFrame()
    for key, values in data.items():
        print(f'values.shape {key} before preprocessing', values.shape)
        values = values.fillna('')
        title = values.iloc[0].to_list()
        if title[0] == '' and title[1] == '':
            values.iloc[:, 0] = values.iloc[:, 0].astype(str) + ' ' + values.iloc[:, 1].astype(str)
            values = values.drop(values.columns[1], axis=1)
            title = values.iloc[0].to_list()

        for i in range(len(title)):
            if i == 0:
                title[i] = 'Комментарии'
            if title[i] == 'Адрес':
                title[i] = 'Улица'
                title[i+1] = 'Дом'
                title[i+2] = 'Квартира'
            if title[i] == 'пол':
                title[i] = 'Пол'

        values = values.set_axis(title, axis='columns')
        values = values.drop(0, axis=0)
        values = values.loc[:, :'Пол']
        values = values.drop_duplicates()
        values = values.set_axis([x for x in range(values.shape[0])], axis='index')

        drop_lines = []
        vib_lines = []
        prib_lines = []
        for i in range(values.shape[0]):
            # выделяем строки, которые надо удалить
            if values.loc[i, 'ФИО'] in ['']:
                drop_lines.append(i)
            elif values.loc[i, 'ФИО'].lower().strip() in month_names:
                drop_lines.append(i)
            elif (str(values.loc[i, 'ДР']).strip()) and (values.loc[i, 'ДР'].date() < min_bd):
                drop_lines.append(i)
            
            # выделяем строки, для прибывших детей
            elif values.loc[i, 'Приб/выб'].lower().strip().startswith('приб'):
                prib_lines.append(i)

            # выделяем строки, для выбывших детей
            elif values.loc[i, 'Приб/выб'].lower().strip().startswith('выб'):
                vib_lines.append(i)
            elif (values.loc[i, 'Орг-ть'] == ''):
                vib_lines.append(i)
            elif (str(values.loc[i, 'ДР']).strip() == ''):
                vib_lines.append(i)

        values = values.set_axis([x for x in range(values.shape[0])], axis='index')

        values['Прибыл'] = pd.Series().fillna('')
        values.loc[prib_lines, 'Прибыл'] = act_year_month
        
        values['Выбыл'] = pd.Series().fillna('')
        values.loc[vib_lines, 'Выбыл'] = act_year_month

        values['ДР'] = values['ДР'].apply(lambda x: '' if isinstance(x, str) else x.strftime('%Y/%m/%d'))

        values = values.drop(drop_lines, axis=0)
        print(f'values.shape {key} after preprocessing', values.shape)

        df = pd.concat([df, values], ignore_index=True)

    print('df.shape', df.shape)
    return df


def make_org(x, org_list):
    if x == '':
        return '', ''
    
    org = str(x).lower().split()

    if len(org) == 1:
        if org[0] in org_list:
            return org[0], ''
        else:
            return '', org[0]
    elif len(org) == 2:
        if org[0] in org_list:
            return org[0], org[1]
        return '', ''.join(org)
    elif len(org) > 2:
        if org[0] in org_list:
            return org[0], ' '.join(org[1:])
        return '', ''.join(org)
    else:
        return '', ''

def add_org(df, org_list):
    print('add_org')
    df['Орг #'] = df['Орг-ть'].apply(lambda x: make_org(x, org_list)[1])
    df['Орг-ть'] = df['Орг-ть'].apply(lambda x: make_org(x, org_list)[0])
    return df

def add_comments(df):
    print('add_comments')
    df['МС'] = pd.Series()
    df['СВО'] = pd.Series()
    df['О'] = pd.Series()
    df['ИНВ'] = pd.Series()
    df['Питание'] = pd.Series()

    for line in range(df.shape[0]):
        comments = df.loc[line, 'Комментарии']
        comments = comments.split()
        if len(comments) > 0:
            for comm in comments:
                if comm.lower() in ['е', 'и', 'см']:
                    df.loc[line, 'Питание'] = comm.lower()
                else:
                    df.loc[line, comm.upper()] = comm.lower()

    df = df.fillna('')
    return df


def check_full_info(df):
    print('check_full_info')
    check_columns = ['ФИО', 'ДР', 'Квартира', 'Улица', 'Дом', 'Пол', 'Орг-ть']
    rows_num = []
    for i in range(df.shape[0]):
        row = df.iloc[i]
        if row.loc['Выбыл'] != '':
            continue
        for j in check_columns:
            if row[j] == '':
                rows_num.append(i)
                break
    return df.iloc[rows_num].sort_values(by='ДР')


def save_df(df, file_full='files/full.xlsx', file_empty='files/empty.xlsx'):
    file_full_list = file_full.split('/')
    file_empty = '/'.join(file_full_list[:-1]) + '/empty.xlsx'

    print(file_full, file_empty)

    print('save_df', df.shape)
    df.to_excel(file_full,
                sheet_name='Список детей',
                index=False,
                header=True,
                startrow=1,
                startcol=0)
    empty_df = check_full_info(df)
    print('save empty_df', empty_df.shape)
    empty_df.to_excel(file_empty,
                sheet_name='Список детей',
                index=False,
                header=True,
                startrow=1,
                startcol=0)
    print('saving - done')
    

def create_new_db(file_path, full_path='files/full.xlsx', today=datetime.now()):
    print('create_new_db')
    today = today.date()
    sheet_names = create_sheet_names(today)
    act_year_month = year_month_now(today)
    min_bd = (today - relativedelta(years=18))
    df = read_data(file_path, sheet_names, min_bd, act_year_month)
    df = add_org(df, general_vars.ORG_OPTIONS.value)
    df['Номер телефона'] = pd.Series()
    df = add_comments(df)
    df = df[general_vars.FILE_COLUMNS_LIST.value]
    save_df(df, full_path)


def create_new_empty_db(file_path):
    empty_df = pd.DataFrame(
        columns=general_vars.FILE_COLUMNS_LIST.value
    )
    save_df(empty_df, file_path)

def find_idx_title_promed(title):
    fio_idx, bd_idx, address_phone_idx, org_idx, prib_idx = 0, 0, 0, 0, 0

    for i in range(len(title)):
        if 'ФИО' in title[i]:
            fio_idx = i
        elif 'Дата рождения' in title[i]:
            bd_idx = i
        elif 'Адрес' in title[i]:
            address_phone_idx = i
        elif 'Посещает образовательное учреждение' in title[i]:
            org_idx = i + 1
        elif 'Прибыл' in title[i]:
            prib_idx = i
        else:
            continue
    
    return fio_idx, bd_idx, address_phone_idx, org_idx, prib_idx

def check_digits_in_list(list):
    return max([i if list[i].isdigit() else -1 for i in range(len(list))]) > -1

def split_address_phone(address_phone_list):
    street, home, appart, phone = '', '', '', ''
    phone = [address_phone_list[-1] if address_phone_list[-1].isdigit() else ''][0]
    home_idx = max([i if 'д.' in address_phone_list[i] or 'д ' in address_phone_list[i] else -1 for i in range(len(address_phone_list))])

    street = address_phone_list[home_idx - 1].capitalize()
    home = address_phone_list[home_idx].split(' ')[-1].lower()

    if phone:
        appart = address_phone_list[-2]
        if address_phone_list[home_idx + 1] != appart:
            home += address_phone_list[home_idx + 1][-1].lower()
        appart = appart.split(' ')[-1]
    else:
        appart = address_phone_list[-1]
        if address_phone_list[home_idx + 1] != appart:
            home += address_phone_list[home_idx + 1][-1].lower()
        appart = appart.split(' ')[-1]

    return street, home, appart, phone

def create_new_promed_db(file_path, full_path='files/full.xlsx', today=datetime.now()):
    doc = load(file_path)

    data = []
    for table in doc.spreadsheet.getElementsByType(Table):
        for row in table.getElementsByType(TableRow):
            row_data = []
            for cell in row.getElementsByType(TableCell):
                val = cell.__str__()
                if '1900-01-00' <= val <= '2000-01-00':
                    val = ''
                row_data.append(val)
            data.append(row_data)
    data = pd.DataFrame(data[26:]).fillna('')

    fio_idx, bd_idx, address_phone_idx, org_idx, prib_idx = find_idx_title_promed(data.iloc[0])

    org_types = general_vars.ORG_TYPES_PROMED.value

    df = pd.DataFrame(
        columns=general_vars.FILE_COLUMNS_LIST.value
    )

    for row_num in range(1, data.shape[0]):
        fio_value = data.iloc[row_num, fio_idx]
        if fio_value != '':
            fio_list = fio_value.split(' ')

            # ячейка ФИО содержит адекватное количество слов и не содержит чисел
            if len(fio_list) >= 2 and not check_digits_in_list(fio_list):
                bd_value = data.iloc[row_num, bd_idx]
                address_phone_value = data.iloc[row_num, address_phone_idx]

                # В строке записаны дата рождения или адрес
                if len(bd_value.split('.')) == 3 or len(address_phone_value.split(' ')) > 1:
                    df.loc[df.shape[0], 'ФИО'] = fio_value #.title()

                    if bd_value != '':
                        df.loc[df.shape[0] - 1, 'ДР'] = pd.to_datetime(bd_value, dayfirst=True).strftime('%Y/%m/%d')
                    
                    address_phone_list = address_phone_value.split(', ')

                    if len(address_phone_list) >= 3:
                        street, home, appart, phone = split_address_phone(address_phone_list)
                        df.loc[df.shape[0] - 1, 'Улица'] = street
                        df.loc[df.shape[0] - 1, 'Дом'] = home
                        df.loc[df.shape[0] - 1, 'Квартира'] = appart
                        df.loc[df.shape[0] - 1, 'Номер телефона'] = str(phone)

                    org_value = data.iloc[row_num, org_idx]
                    for key, value in org_types.items():
                        if value[0].search(org_value):
                            df.loc[df.shape[0] - 1, 'Орг-ть'] = value[1]
                            if value[2]:
                                df.loc[df.shape[0] - 1, 'Орг #'] = org_value.upper()

                    prib_value = data.iloc[row_num, prib_idx]
                    if prib_value != '':
                        df.loc[df.shape[0] - 1, 'Прибыл'] = pd.to_datetime(prib_value, dayfirst=True).strftime('%Y/%m')
    save_df(df, full_path)


if __name__ == '__main__':
    test_path = r'C:\Users\valer\Documents\course_work_3\ПЕРЕПИСЬ 12 уч. март 25Г.xlsx'
    create_new_db(test_path, datetime(2025, 3, 30))