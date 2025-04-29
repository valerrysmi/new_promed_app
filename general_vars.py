from enum import Enum
import re

class general_vars(Enum):
    WORK_WITH_DB = None
    STREET_OPTIONS = []
    GENDER_OPTIONS = ["м", "ж"]
    ORG_OPTIONS = ["н/о", "дс", "шк", "суз", "вуз", "раб"]
    FULL_FILE_PATH = 'files/full.xlsx'
    REPORTS_FILE_PATH = 'files/reports.xlsx'
    FILE_COLUMNS_LIST = ['ФИО', 'ДР', 
                        'Улица', 'Дом', 'Квартира',
                        'Пол', 'Орг-ть', 'Орг #', 'Номер телефона',
                        'Комментарии', 'Прибыл', 'Выбыл',
                        'МС', 'СВО', 'О', 'ИНВ', 'Питание']
    ORG_TYPES_PROMED = {
        0: (re.compile(r'неорганизованный', re.IGNORECASE), 'н/о', False),
        1: (re.compile(r'сад|дс|д/с|мдоу', re.IGNORECASE), 'дс', True),
        2: (re.compile(r'мбоу|сош|нош|школа|гимназия|цо|центр|амтэк', re.IGNORECASE), 'шк', True),
        3: (re.compile(r'колледж|техникум|члмт|чтэк', re.IGNORECASE), 'суз', True),
        4: (re.compile(r'университет', re.IGNORECASE), 'вуз', True),
        5: (re.compile(r'работает', re.IGNORECASE), 'раб', False)
    }
    PROTECTED_COLUMNS = ['МС', 'СВО', 'О', 'ИНВ', 'Питание']
    GENERAL_COLUMNS = ['ФИО', 'ДР', 'Улица', 'Дом', 'Квартира', 'Пол', 'Орг-ть', 'Орг #', 'Комментарии']
    REPORT_NAMES = ['Возраст', 'Орг-ть всего', 'Орг-ть - Возраст', 'Юноши 18', 'По домам',  'По домам кол-во', 
                    'МС список', 'МС кол-во', 'СВО', 'О', 'Инв', 'Неорг', 
                    'План по', 'Приб-Выб', 'Прибыл список', 'Выбыл список']

general_vars_dict = {
    "STREET_OPTIONS" : [],
}
