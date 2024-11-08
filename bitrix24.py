""" Этот код представляет собой пример использования Python для взаимодействия с Bitrix24 API.
Основная цель скрипта заключается в поиске кандидата в базе по id, сохранения данных в файл 
и прикрипления ссылки на файл к кандидату.
Также возможность создания смарт процесса для Bitrix24, которая получает необохдимые данные при
загрузки из файла. 

Импортированные модули:
- openpyxl: Модуль для работы с файлами Excel (.xlsx). 
- requests: Модуль для для выполнения HTTP-запросов.
- datetime: Модуль предоставляет классы для работы с датами и временем.
- logging: Модуль для логирования. 
"""


import openpyxl
from openpyxl.styles import Font, Alignment
import requests
from datetime import datetime
import logging


# Настройка логгирования. Данные хранятся в файле youtube.log с уровнем логирования INFO. 
logging.basicConfig(
    filename = 'bitrix24.log',  
    level = logging.INFO,  
    format = '%(asctime)s - %(levelname)s - %(message)s', 
)


# Настройки констант для Bitrix24 API 
# (WEBHOOK - должен предоставлять доступ к определенным модулям)
BITRIX_WEBHOOK_URL = 'https://your_domain.bitrix24.ru/rest/1/your_webhook/'


def get_candidate_data(candidate_id:int) -> dict:
    """Функция get_candidate_data получает данные кандидата из системы Bitrix24
    по уникальному идентификатору.
    
    :param candidate_id: уникальный идентификатор кандидата.
    :return: словарь с информацией о кандидате.
    """
    
    logging.info(f'start get_candidate_data - {candidate_id}')
    
    url = f"{BITRIX_WEBHOOK_URL}crm.lead.get" # crm для получения данных
    params = {
        "id": candidate_id
    }
    response = requests.get(url, params=params) # запрос для получения данных
    try:
        logging.info(f'candidate details received')
        return response.json() 
    except requests.exceptions.HTTPError as e:
        logging.error("HTTP error occurred: %s", e)
        return None
    except ValueError:
        logging.error("Error decoding JSON. Response:", response.text)
        return None


def save_candidate_to_excel(candidate_data:dict, file_name:str='candidates') ->str:
    """Функция предназначена для сохранения данных о кандидате в файл 
    формата Excel (*.xlsx*).
    
    :param candidate_data: словарь с данными кандидата, который будет записан в Excel-файл.
    :param filename (необязательный параметр): имя файла, куда будут сохраняться 
    данные. По умолчанию используется значение `'candidates'`.
    return: str о том, что данные были успешно сохранены в указанный файл."""
    
    try: 
        logging.info(f'start save_candidate_to_excel')
        # Создание новой рабочей книги Excel и получение активного листа
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Кандидаты'
    
        # Стили заголовков в создаваемом файле
        header_font = Font(bold=True)
        align_center = Alignment(horizontal='center')
    
        # Заголовки столбцов
        headers = ['ID', 'Имя', 'Фамилия', 'Телефон', 'Email', 'Дата создания']
        sheet.append(headers)
        for col, value in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col)
            cell.value = value
            cell.font = header_font
            cell.alignment = align_center
    
        # Заполнение данных по столбцам
        row_num = 2
        sheet.cell(row=row_num, column=1).value = candidate_data['result']['ID']
        sheet.cell(row=row_num, column=2).value = candidate_data['result']['NAME']
        sheet.cell(row=row_num, column=3).value = candidate_data['result']['LAST_NAME']
        sheet.cell(row=row_num, column=4).value = candidate_data['result']['PHONE'][0]['VALUE'] if len(candidate_data['result']['PHONE'][0]['VALUE']) > 0 else ''
        sheet.cell(row=row_num, column=5).value = candidate_data['result']['EMAIL'][0]['VALUE'] if len(candidate_data['result']['EMAIL'][0]['VALUE']) > 0 else ''
        sheet.cell(row=row_num, column=6).value = datetime.strptime(candidate_data['result']['DATE_CREATE'], '%Y-%m-%dT%H:%M:%S%z').strftime('%d.%m.%Y')
        row_num += 1
        
        # названия созданного файла с указанием времени создания
        file_name_save = f"{file_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        workbook.save(file_name_save)
        logging.info(f'Candidate data saved in {file_name_save}')
        return f"Данные кандидата сохранены в {file_name_save}"
    except Exception as e:
        logging.error("Error save: %s", e)
        return None


def upload_file_to_lead(file_name:str) -> int:
    """Функция предназначена создания поля в карточке кандидата Bitrix24 
    для прикрипления файла.

    :param filename: имя файла, который необходимо прикрепить к карточке
    :return: id созданного поля.
    """
    
    logging.info(f'Start upload_file_to_lead')
    
    url = f'{BITRIX_WEBHOOK_URL}crm.lead.userfield.add' # crm для создания поля
    try:
        payload = {
            "fields": {
                "FIELD_NAME": "LINK_TO_CANDIDATS",
                "EDIT_FORM_LABEL": f"Ссылка на файл {file_name.split('_')[0]}",
                "LIST_COLUMN_LABEL": f"Ссылка на файл '{file_name.split('_')[0]}'",
                "USER_TYPE_ID": "file",
                "MULTIPLE": "N",
                "MANDATORY": "N",
                "SHOW_FILTER": "N",
                "SHOW_IN_LIST": "Y",
                "IS_SEARCHABLE": "N",
                "SORT": 100,
                "XML_ID": "LINK_TO_CANDIDATE_FILE"
            }
        }
        response = requests.post(url, json=payload) # запрос для создания поля
        result = response.json()
        
        logging.info(f"Поле успешно создано с ID: {result['result']}")
        return result['result']
    except Exception as e:
        logging.error(f"Ошибка при создании поля: {e}")
        

def save_link_to_file(field_id:int, file_path:str, candidate_id:int):
    """Функция предназначена для прикрипления ссылки на файла с данными кандидата
    к карточке кандидата Bitrix24.

    :param field_id: id поля для прикрепления ссылки.
    :param file_path: путь к файлу.
    :param candidate_id: id кандидата к которому прикрепить файл.
    """
    
    logging.info(f'start save_link_to_file')
    url = f'{BITRIX_WEBHOOK_URL}crm.lead.update.json' # crm для создания поля
    try: 
        payload = {
            "id": candidate_id,
            "fields": {
                field_id: {"value": file_path}
            }
        }

        response = requests.post(url, json=payload)  # запрос для создания поля
        result = response.json()
        logging.info(f"Ссылка на файл успешно сохранена у кандидата - {candidate_id}")
    except Exception as e:
        logging.error(f"Ошибка при сохранении ссылки на файл: {e}")


def read_from_excel(file_name:str) ->list:
    """Функция предназначена чтения файла формата Excel (*.xlsx*)
    
    :param filename: имя файла, который необходимо открыть
    :return: list с данными из файла. """
    
    logging.info(f'start read_from_excel')
    try:
        # Открытие книги
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active
        
        data = []
        headers = next(ws.rows, None)  # Пропустить первую строку с заголовками
        for row in ws.iter_rows(min_row=2, values_only=True):
            data.append({
                'TITLE': row[0], 
                'NAME': row[1],  
                'LAST_NAME': row[2],  
                'PHONE': row[3],  
                'EMAIL': row[4]
            })
        logging.info(f'File read successfully - {file_name}')
        return data
    except Exception as e:
        logging.error("Error save: %s", e)
        

def create_smart_process(data:list):
    """Функция предназначена создания смарт процесса в Bitrix24
    :param data: данные для загрузки смарт процесса (list). """
    
    logging.info(f'start create_smart_process')
    
    try: 
        smart_process_url = f'{BITRIX_WEBHOOK_URL}crm.lead.add' # crm для создания процесса
        for item in data:
            lead_data = {
                'fields': {
                    'TITLE': item['TITLE'],
                    'NAME': item['NAME'] if item['NAME'] else 'Empty name',
                    'LAST_NAME': item['LAST_NAME'],
                    'PHONE': [{'VALUE': item['PHONE'], 'VALUE_TYPE': 'HOME'}] if item['PHONE'] else [],
                    'EMAIL': [{'VALUE': item['EMAIL'], 'VALUE_TYPE': 'HOME'}] if item['EMAIL'] else []
                    }
                }
            response = requests.post(smart_process_url, json=lead_data)
            if response.status_code == 200:
                logging.info(f"Смарт-процесс '{item['TITLE']}' успешно создан!")
            else:
                logging.error(f"Ошибка при создании смарт-процесса: {response.text}")
    except Exception as e:
        logging.error("Error save: %s", e)

      
def main_candidate_data():
    """ Основная функция которая получает информацию о id кандидата, производит выгрузку данных в
    функции get_candidate_data. 
    Резлуьтат записывается через функцию save_candidate_to_excel в файл формата .xlsx и ссылка 
    прикрепляется к карточки кандидата - save_link_to_file/
    """
    candidate_id = int(input("Напишите id кандидата: "))
    
    # получаем данные по кандидату
    candidate_data = get_candidate_data(candidate_id)
    if candidate_data:
        # генерируем Excel-файл
        excel_file = save_candidate_to_excel(candidate_data)
        print(f'Создание Excel-файла {excel_file} завершено!')
    
    # создания поля для прикрипления ссылки
    field_id = upload_file_to_lead('candidate_4_20241108_1700.xlsx')
    if field_id:
        save_link_to_file(field_id, 'candidate_4_20241108_1700.xlsx', candidate_id)
        print(f'Прикрепление карточки к кандидату завершено!')
        
 
def main_smart_process(): 
    """ Основная функция которая получает информацию файле формата .xlsx, который открывается для 
    чтения через функцию read_from_excel и из полученного результата создает смарт процесс
    с помощью функции create_smart_process 
    """
    data = read_from_excel('test_crm.xlsx')
    if data:
        create_smart_process(data)
        print('Добавление смарт процессов завершено!')


if __name__ == '__main__':
    main_candidate_data()
    main_smart_process()
