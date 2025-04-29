from general_vars import general_vars
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

def open_db(act_sheet_name):
    df = pd.read_excel(general_vars.REPORTS_FILE_PATH.value, sheet_name=act_sheet_name, header=0)
    df = df.dropna(how='all').fillna('')
    return df

def get_age(birthday, day_for_report):
    birthday = pd.to_datetime(birthday).date()
    for age in range(20):
        date_age = (day_for_report - relativedelta(years=age)) #.date()
        if birthday > date_age:
            return age - 1
        elif birthday == date_age:
            return age

def date_i_month(next_month, next_month_year, i, day_for_report):
    date_i_m_month = (next_month - i) % 12 + 12 * ((next_month - i) % 12 == 0)
    date_i_m_year = next_month_year - 1 * (date_i_m_month > day_for_report.month) - ((i - 1) // 12)
    return str(date_i_m_year) + '/' + ('0' + str(date_i_m_month))[-2:]

def prib_vib_age(x):
    x = x.split('!')

    if x[1] == '': # выбывшие, у которых не указана ДР
        return 0

    pd_y, pd_m = map(int, x[0].split('/'))
    bd_y, bd_m = map(int, x[1].split('/'))
    age = pd_y - bd_y - 1 * (pd_m < bd_m)
    return age 

def make_reports(day_for_report=None):
    input_file = general_vars.FULL_FILE_PATH.value
    result_file = general_vars.REPORTS_FILE_PATH.value

    if day_for_report is None:
        day_for_report = datetime.today().date()
    
    full_df = pd.read_excel(input_file)
    full_df = full_df.set_axis(full_df.iloc[0], axis=1).drop(0, axis=0).fillna('')
    full_df['ДР'] = full_df['ДР'].apply(lambda x: x if len(x.split('/')) == 3 else day_for_report.strftime('%Y/%m/%d'))

    act_full_df = full_df[full_df['Выбыл'] == '']

    org_list = general_vars.ORG_OPTIONS.value
    general_info = general_vars.GENERAL_COLUMNS.value
    address = ['Улица', 'Дом', 'Квартира']
    ages = pd.DataFrame(np.array(list(map(str, np.arange(18))) + ['']), columns=['Возраст'])

    act_full_df['Возраст'] = act_full_df['ДР'].apply(lambda x: str(get_age(x, day_for_report)))

    age_stat = (
        act_full_df[['ФИО', 'Возраст']]
        .groupby('Возраст')
        .count()
        .rename(columns={'ФИО': 'Кол-во'})
        .reset_index()
    )
    age_stat = (
        ages[:-1]
        .merge(age_stat, how='left', left_on='Возраст', right_on='Возраст')
        .drop(columns=['Возраст'])
        .fillna(0)
        .astype(int)
    )
    age_stat.loc['0-18'] = age_stat.iloc[0:18].sum()['Кол-во']
    age_stat.loc['0-15'] = age_stat.iloc[0:15].sum()['Кол-во']
    age_stat.loc['0-3'] = age_stat.iloc[0:3].sum()['Кол-во']
    age_stat.loc['3-7'] = age_stat.iloc[3:7].sum()['Кол-во']
    age_stat.loc['7-15'] = age_stat.iloc[7:15].sum()['Кол-во']

    age_stat = (
        age_stat
        .reset_index()
        .rename(columns={'index' : 'Возраст'})
    )

    custom_order = {}
    for i, org_type in enumerate(org_list):
        custom_order[org_type] = i

    org_stat = (
        act_full_df[act_full_df['Возраст'] != '18'][['ФИО', 'Орг-ть']]
        .groupby('Орг-ть')
        .count()
        .sort_index(key=lambda x: x.map(custom_order))
        .reset_index()
        .rename(columns={'ФИО' : 'Кол-во'})
    )
    org_stat.loc['Сумма', 'Орг-ть'] = 'Сумма'
    org_stat.loc['Сумма', 'Кол-во'] = org_stat.iloc[0:5]['Кол-во'].sum()

    org_age_stat = (
        pd.pivot_table(
            (
                act_full_df[['ФИО', 'Орг-ть', 'Возраст']]
                .groupby(['Возраст', 'Орг-ть'])
                .count()
                .reset_index()
            ), 
            values='ФИО', columns='Орг-ть', index='Возраст'
        ).fillna(0)
        .astype(int)
        .sort_index(axis=1, key=lambda x: x.map(custom_order))
        .reset_index()
    )
    org_age_stat = (
        ages
        .merge(org_age_stat, how='left', left_on='Возраст', right_on='Возраст')
        .drop(columns=['Возраст'])
        .fillna(0)
        .astype(int)
    )

    org_age_stat.loc['0-18'] = org_age_stat.iloc[0:18].sum()
    org_age_stat.loc['0-15'] = org_age_stat.iloc[0:15].sum()
    org_age_stat.loc['0-3'] = org_age_stat.iloc[0:3].sum()
    org_age_stat.loc['3-7'] = org_age_stat.iloc[3:7].sum()
    org_age_stat.loc['7-15'] = org_age_stat.iloc[7:15].sum()
    org_age_stat.loc[:, 'Сумма'] = org_age_stat.sum(axis=1)

    org_age_stat = (
        org_age_stat
        .reset_index()
        .rename(columns={'index' : 'Возраст'})
    )

    year_18 = str(day_for_report.year - 18)
    man_18_df = act_full_df[(act_full_df['Пол'] == 'м') & (act_full_df['ДР'].apply(lambda x: x.split('/')[0] == year_18))]
    man_18_df = man_18_df[general_info]
    man_18_df['ДР'] = pd.to_datetime(man_18_df['ДР']).dt.strftime('%d.%m.%Y')

    home_df = act_full_df.sort_values(by=address)[general_info]
    home_df['ДР'] = pd.to_datetime(home_df['ДР']).dt.strftime('%d.%m.%Y')

    home_df_stat = (
        home_df[address]
        .groupby(by=address[:-1])
        .count()
        .reset_index()
        .rename(columns={'Квартира' : 'Кол-во'})
    )

    multi_children_df_full = (
        act_full_df[act_full_df['МС'] == 'мс']
        .sort_values(by = address + ['ДР'])
    )

    multi_children_df = multi_children_df_full[general_info]
    multi_children_df['ДР'] = pd.to_datetime(multi_children_df['ДР']).dt.strftime('%d.%m.%Y')

    multi_children_stat_df = pd.DataFrame()

    multi_children_stat_df.loc['Всего семей', 'Количество'] = (
        multi_children_df[address]
        .T
        .apply(lambda x: ' '.join(map(str, x)))
        .nunique()
    )

    multi_children_stat_df.loc['Всего детей', 'Количество'] = multi_children_df['ФИО'].nunique()

    multi_children_stat_df.loc['До 1г', 'Количество'] = (
        multi_children_df_full[multi_children_df_full['Возраст'].apply(lambda x: int(x) == 0)]['ФИО']
        .nunique()
    )

    multi_children_stat_df.loc['До 3л', 'Количество'] = (
        multi_children_df_full[multi_children_df_full['Возраст'].apply(lambda x: int(x) < 3)]['ФИО']
        .nunique()
    )

    multi_children_stat_df.loc['До 6л', 'Количество'] = (
        multi_children_df_full[multi_children_df_full['Возраст'].apply(lambda x: int(x) < 6)]['ФИО']
        .nunique()
    )

    multi_children_stat_df = multi_children_stat_df.astype(int)

    multi_children_stat_df = (
        multi_children_stat_df
        .reset_index()
        .rename(columns={'index' : 'МС'})
    )

    svo_df_full = (
        act_full_df[act_full_df['СВО'] == 'сво']
        .sort_values(by=address)
    )

    svo_df = svo_df_full[general_info]
    svo_df['ДР'] = pd.to_datetime(svo_df['ДР']).dt.strftime('%d.%m.%Y')

    opeka_df_full = (
        act_full_df[act_full_df['О'] == 'о']
    )

    opeka_df = opeka_df_full[general_info]
    opeka_df['ДР'] = pd.to_datetime(opeka_df['ДР']).dt.strftime('%d.%m.%Y')

    inv_df_full = (
        act_full_df[act_full_df['ИНВ'] == 'инв']
    )

    inv_df = inv_df_full[general_info]
    inv_df['ДР'] = pd.to_datetime(inv_df['ДР']).dt.strftime('%d.%m.%Y')

    neorg_df_full = act_full_df[act_full_df['Орг-ть'] == 'н/о'][general_info]
    neorg_df = neorg_df_full.copy()
    neorg_df['ДР'] = pd.to_datetime(neorg_df['ДР']).dt.strftime('%d.%m.%Y')


    next_month = (day_for_report.month + 1) % 12
    next_month_year = day_for_report.year + 1 * (next_month == 1)

    plan_po = pd.DataFrame()
    empty_line = pd.DataFrame(np.array([''] * neorg_df_full.shape[1]), index=neorg_df_full.columns).T

    for i in range(1, 13):
        date_i_m = date_i_month(next_month, next_month_year, i, day_for_report)
        plan_po_i = neorg_df_full[neorg_df_full['ДР'].apply(lambda x: x[:7] == date_i_m)]
        if plan_po_i.shape[0] > 0:
            title = empty_line.copy()
            title.iloc[0, 0] = f'{i} мес'
            plan_po_i.loc[:, 'ДР'] = pd.to_datetime(plan_po_i.loc[:, 'ДР']).dt.strftime('%d.%m.%Y')
            plan_po = pd.concat([plan_po, empty_line, title, plan_po_i])

    for i in [1 * 12 + 3, 1 * 12 + 6, 1 * 12 + 9, 2 * 12, 2 * 12 + 6]:
        date_i_m = date_i_month(next_month, next_month_year, i, day_for_report)
        plan_po_i = neorg_df_full[neorg_df_full['ДР'].apply(lambda x: x[:7] == date_i_m)]
        if plan_po_i.shape[0] > 0:
            title = empty_line.copy()
            title.iloc[0, 0] = f'{i // 12} г {i % 12} мес'
            plan_po_i['ДР'] = pd.to_datetime(plan_po_i['ДР']).dt.strftime('%d.%m.%Y')
            plan_po = pd.concat([plan_po, empty_line, title, plan_po_i])

    for i in range(3, 18):
        date_i_m = date_i_month(next_month, next_month_year, i * 12, day_for_report)
        plan_po_i = neorg_df_full[neorg_df_full['ДР'].apply(lambda x: x[:7] == date_i_m)]
        if plan_po_i.shape[0] > 0:
            title = empty_line.copy()
            title.iloc[0, 0] = f'{i} г'
            plan_po_i['ДР'] = pd.to_datetime(plan_po_i['ДР']).dt.strftime('%d.%m.%Y')
            plan_po = pd.concat([plan_po, empty_line, title, plan_po_i])

    prib_df = (
        full_df[['ФИО', 'Прибыл']]
        .groupby(['Прибыл'])
        .count()
        .rename(columns={'ФИО': 'Кол-во'})
        .reset_index()
        .rename(columns={'Прибыл' : 'Дата'})
    )

    vib_df = (
        full_df[['ФИО', 'Выбыл']]
        .groupby(['Выбыл'])
        .count()
        .rename(columns={'ФИО': 'Кол-во'})
        .reset_index()
        .rename(columns={'Выбыл' : 'Дата'})
    )

    prib_vib_df = (
        prib_df
        .merge(vib_df, how='outer', on='Дата', suffixes=(' Приб', ' Выб'))
        .fillna(0)
    )
    prib_vib_df['Кол-во Приб'] = prib_vib_df['Кол-во Приб'].astype(int)
    prib_vib_df['Кол-во Выб'] = prib_vib_df['Кол-во Выб'].astype(int)

    full_df['ДР_гм'] = full_df['ДР'].apply(lambda x: x[:-3])
    full_df['Приб_ДР'] = full_df['Прибыл'] + '!' + full_df['ДР_гм']
    full_df['Возраст_Прибыл'] = full_df['Приб_ДР'].apply(lambda x: '' if x.split('!')[0] == '' else prib_vib_age(x))

    prib_age_df = (
        pd.pivot_table(
            (
            full_df[['ФИО', 'Прибыл', 'Возраст_Прибыл']]
            .groupby(['Прибыл', 'Возраст_Прибыл'])
            .count()
            .rename(columns={'ФИО': 'Кол-во'})
            .reset_index()
            .rename(columns={'Прибыл' : 'Поступило'})
            .reset_index()
            ),
            values='Кол-во', index='Возраст_Прибыл', columns='Поступило'
        )
        .fillna(0)
        .astype(int)
    )

    prib_age_df = (
        ages
        .merge(prib_age_df, how='left', left_on='Возраст', right_on='Возраст_Прибыл')
        .drop(columns=['Возраст', ''])
        .fillna(0)
        .astype(int)
    )

    prib_age_stat_df = pd.DataFrame()
    prib_age_stat_df['до 18 лет'] = prib_age_df.iloc[0:18].sum()
    prib_age_stat_df['до 15 лет'] = prib_age_df.iloc[0:15].sum()
    prib_age_stat_df['до года'] = prib_age_df.iloc[0:1].sum()
    prib_age_stat_df['новорож.'] = prib_age_df.iloc[0]
    prib_age_stat_df['1 - 2 лет'] = prib_age_df.iloc[1:2].sum()
    prib_age_stat_df['2 - 3 года'] = prib_age_df.iloc[2:3].sum()
    prib_age_stat_df['3 - 7 лет'] = prib_age_df.iloc[3:7].sum()
    prib_age_stat_df['7 - 15 лет'] = prib_age_df.iloc[7:15].sum()
    prib_age_stat_df['15 - 16'] = prib_age_df.iloc[15:16].sum()
    prib_age_stat_df['16 - 17'] = prib_age_df.iloc[16:17].sum()
    prib_age_stat_df['17 - 18'] = prib_age_df.iloc[17:18].sum()
    prib_age_stat_df = prib_age_stat_df.T

    prib_age_stat_df = (
        prib_age_stat_df
        .reset_index()
        .rename(columns={'index': 'Возраст'})
    )
    prib_age_stat_df = (
        pd.concat([pd.DataFrame(data={'Категория' : ['Прибыло']}), prib_age_stat_df])
        .fillna('')
    )
    prib_age_stat_df.iloc[4, 0] = 'Из них'

    full_df['Выб_ДР'] = full_df['Выбыл'] + '!' + full_df['ДР_гм']
    full_df['Возраст_Выбыл'] = full_df['Выб_ДР'].apply(lambda x: '' if x.split('!')[0] == '' else prib_vib_age(x))

    vib_age_df = (
        pd.pivot_table(
            (
            full_df[['ФИО', 'Выбыл', 'Возраст_Выбыл']]
            .groupby(['Выбыл', 'Возраст_Выбыл'])
            .count()
            .rename(columns={'ФИО': 'Кол-во'})
            .reset_index()
            .rename(columns={'Выбыл' : 'Выбыло'})
            .reset_index()
            ),
            values='Кол-во', index='Возраст_Выбыл', columns='Выбыло'
        )
        .fillna(0)
        .astype(int)
        
    )

    vib_age_df = (
        pd.concat([ages, vib_age_df])
        .drop(columns=['Возраст', ''])
        .fillna(0)
        .astype(int)
    )

    vib_age_stat_df = pd.DataFrame()
    vib_age_stat_df['до 18 лет'] = vib_age_df.iloc[0:18].sum()
    vib_age_stat_df['до 15 лет'] = vib_age_df.iloc[0:15].sum()
    vib_age_stat_df['до года'] = vib_age_df.iloc[0:1].sum()
    vib_age_stat_df['новорож.'] = vib_age_df.iloc[0]
    vib_age_stat_df['1 - 2 лет'] = vib_age_df.iloc[1:2].sum()
    vib_age_stat_df['2 - 3 года'] = vib_age_df.iloc[2:3].sum()
    vib_age_stat_df['3 - 7 лет'] = vib_age_df.iloc[3:7].sum()
    vib_age_stat_df['7 - 15 лет'] = vib_age_df.iloc[7:15].sum()
    vib_age_stat_df['15 - 16'] = vib_age_df.iloc[15:16].sum()
    vib_age_stat_df['16 - 17'] = vib_age_df.iloc[16:17].sum()
    vib_age_stat_df['17 - 18'] = vib_age_df.iloc[17:18].sum()
    vib_age_stat_df = vib_age_stat_df.T

    vib_age_stat_df = (
        vib_age_stat_df
        .reset_index()
        .rename(columns={'index': 'Возраст'})
    )
    vib_age_stat_df = (
        pd.concat([pd.DataFrame(data={'Категория' : ['Выбыло']}), vib_age_stat_df])
        .fillna('')
    )
    vib_age_stat_df.iloc[4, 0] = 'Из них'

    prib_vib_age_stat = (
        pd.concat([prib_age_stat_df, vib_age_stat_df])
        .fillna(0)
    )

    prib_months_data = (
        full_df[full_df['Прибыл'] != '']
        .sort_values(by='Прибыл', ascending=True)
    )

    prib_months = pd.DataFrame(columns=prib_months_data.columns)

    act_month = ''
    for line_num in range(prib_months_data.shape[0]):
        line = prib_months_data.iloc[line_num]
        if line['Прибыл'] != act_month:
            act_month = line.loc['Прибыл']
            prib_months.loc[prib_months.shape[0], 'ФИО'] = ''
            prib_months.loc[prib_months.shape[0], 'ФИО'] = act_month
        
        prib_months.loc[prib_months.shape[0], :] = line

    prib_months = prib_months.fillna('')
    prib_months = prib_months[general_info]

    vib_months_data = (
        full_df[full_df['Выбыл'] != '']
        .sort_values(by='Выбыл', ascending=True)
    )

    vib_months = pd.DataFrame(columns=vib_months_data.columns)

    act_month = ''
    for line_num in range(vib_months_data.shape[0]):
        line = vib_months_data.iloc[line_num]
        if line['Выбыл'] != act_month:
            act_month = line.loc['Выбыл']
            vib_months.loc[vib_months.shape[0], 'ФИО'] = ''
            vib_months.loc[vib_months.shape[0], 'ФИО'] = act_month
        
        vib_months.loc[vib_months.shape[0], :] = line

    vib_months = vib_months.fillna('')
    vib_months = vib_months[general_info]

    with pd.ExcelWriter(result_file, engine='openpyxl') as writer:
        age_stat.to_excel(writer, sheet_name='Возраст', index=False)
        org_stat.to_excel(writer, sheet_name='Орг-ть всего', index=False)
        org_age_stat.to_excel(writer, sheet_name='Орг-ть - Возраст', index=False)
        man_18_df.to_excel(writer, sheet_name='Юноши 18', index=False)
        home_df.to_excel(writer, sheet_name='По домам', index=False)
        home_df_stat.to_excel(writer, sheet_name='По домам кол-во', index=False)
        multi_children_df.to_excel(writer, sheet_name='МС список', index=False)
        multi_children_stat_df.to_excel(writer, sheet_name='МС кол-во', index=False)
        svo_df.to_excel(writer, sheet_name='СВО', index=False)
        opeka_df.to_excel(writer, sheet_name='О', index=False)
        inv_df.to_excel(writer, sheet_name='Инв', index=False)
        neorg_df.to_excel(writer, sheet_name='Неорг', index=False)
        plan_po.to_excel(writer, sheet_name='План по', index=False)
        prib_vib_age_stat.to_excel(writer, sheet_name='Приб-Выб', index=False)
        prib_months.to_excel(writer, sheet_name='Прибыл список', index=False)
        vib_months.to_excel(writer, sheet_name='Выбыл список', index=False)




