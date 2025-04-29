import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from tkcalendar import DateEntry
import re

from general_vars import general_vars, general_vars_dict

def open_db():
    df = load_and_prepare_data()
    return df

def format_phone_number(value):
    if pd.isna(value) or value == '':
        return ''
    phone_str = str(value)
    if phone_str.endswith('.0'):
        phone_str = phone_str[:-2]
    return phone_str

def format_apartment_number(value):
    if pd.isna(value) or value == '':
        return ''
    return str(int(value))

def parse_and_format_date(date_value):
    if pd.isna(date_value) or date_value == '':
        return ''
    if isinstance(date_value, str) and re.match(r'\d{2}\.\d{2}\.\d{4}', date_value):
        return date_value
    return datetime.strptime(str(date_value), '%Y/%m/%d').strftime('%d.%m.%Y')    

def load_and_prepare_data():
    print('load_and_prepare_data')
    df = pd.read_excel(general_vars.FULL_FILE_PATH.value, sheet_name='Список детей', header=1)
    df = df.dropna(how='all').fillna('')

    existing_streets = df['Улица'].unique()
    general_vars_dict['STREET_OPTIONS'] = sorted(list(set(general_vars_dict['STREET_OPTIONS'] + [x for x in existing_streets if x])))
    
    df['ДР'] = df['ДР'].apply(parse_and_format_date)
    df['Квартира'] = df['Квартира'].apply(format_apartment_number)
    df['Номер телефона'] = df['Номер телефона'].apply(format_phone_number)

    return df

def save_new_record(entries):
    new_data = {}

    for col in general_vars.PROTECTED_COLUMNS.value:
        new_data[col] = ''
    
    if 'Комментарии' in entries:
        comments = entries['Комментарии'].get()
        new_data['Комментарии'] = comments
        
        comments = str(comments).split()
        if len(comments) > 0:
            for comm in comments:
                comm_lower = comm.lower()
                if comm_lower in ['е', 'и', 'см']:
                    new_data['Питание'] = comm_lower
                else:
                    comm_upper = comm.upper()
                    if comm_upper in general_vars.FILE_COLUMNS_LIST.value:
                        new_data[comm_upper] = comm_lower

    for col in [c for c in general_vars.FILE_COLUMNS_LIST.value if c not in general_vars.PROTECTED_COLUMNS.value and c != 'Комментарии']:
        if col in ['ДР', 'Прибыл', 'Выбыл']:
            date_str = entries[col].get()
            if date_str:
                try:
                    date_obj = datetime.strptime(date_str, '%d.%m/%Y')
                    if col == 'ДР':
                        new_data[col] = date_obj.strftime('%d.%m.%Y')
                    else:
                        new_data[col] = date_obj.strftime('%Y/%m')
                except ValueError:
                    new_data[col] = ''
            else:
                new_data[col] = ''
        elif col in ['Улица', 'Пол', 'Орг-ть']:
            new_data[col] = entries[col].get() if hasattr(entries[col], 'get') else str(entries[col])
        else:
            value = entries[col].get() if isinstance(entries[col], (ttk.Entry, tk.Entry)) else str(entries[col])
            if col == 'Квартира':
                try:
                    new_data[col] = int(value) if value.isdigit() else value
                except ValueError:
                    new_data[col] = value
            elif col == 'Номер телефона':
                new_data[col] = format_phone_number(value)
            else:
                new_data[col] = value

    if new_data['Улица'] not in general_vars_dict['STREET_OPTIONS']:
        general_vars_dict['STREET_OPTIONS'] = sorted(general_vars_dict['STREET_OPTIONS'].append(new_data['Улица']))

    return new_data

def save_changes(df):
    try:
        save_df = df.copy()
        
        if 'ДР' in save_df.columns:
            save_df['ДР'] = save_df['ДР'].apply(
                lambda x: datetime.strptime(x, '%d.%m.%Y').strftime('%Y/%m/%d') 
                if x and re.match(r'\d{2}\.\d{2}\.\d{4}', str(x)) 
                else x
            )
        
        if 'Прибыл' in save_df.columns or 'Выбыл' in save_df.columns:
            for col in ['Прибыл', 'Выбыл']:
                if col in save_df.columns:
                    save_df[col] = save_df[col].apply(
                        lambda x: datetime.strptime(x, '%Y/%m').strftime('%Y/%m') 
                        if x and re.match(r'\d{4}/\d{2}', str(x)) 
                        else x
                    )
        
        with pd.ExcelWriter(general_vars.FULL_FILE_PATH, engine='openpyxl') as writer:
            save_df.to_excel(
                writer,
                sheet_name='Список детей',
                index=False,
                header=True,
                startrow=1,
                startcol=0
            )
        
        messagebox.showinfo("Успех", "Изменения сохранены в файл")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")



