import unittest
from unittest.mock import patch, MagicMock

from bitrix24 import (
    BITRIX_WEBHOOK_URL,
    get_candidate_data,
    save_candidate_to_excel,
    upload_file_to_lead, 
    save_link_to_file, 
    create_smart_process, 
    read_from_excel,
)


class TestBitrix24(unittest.TestCase):

    @patch('requests.get')
    def test_get_candidate_data(self, mock_requests_get):
        # Подготавливаем фиктивный ответ от API
        mock_response = MagicMock()
        mock_response.json.return_value = {'result': {'ID': 123, 'NAME': 'Иван', 'LAST_NAME': 'Иванов', 'PHONE': [{'VALUE': '+79991234567'}], 'EMAIL': [{'VALUE': 'ivanov@example.com'}], 'DATE_CREATE': '2023-10-01T12:00:00+0300'}}
        mock_requests_get.return_value = mock_response

        # Вызываем функцию
        result = get_candidate_data(123)

        # Проверяем результат
        self.assertEqual(result, {'result': {'ID': 123, 'NAME': 'Иван', 'LAST_NAME': 'Иванов', 'PHONE': [{'VALUE': '+79991234567'}], 'EMAIL': [{'VALUE': 'ivanov@example.com'}], 'DATE_CREATE': '2023-10-01T12:00:00+0300'}})

    @patch('openpyxl.Workbook')
    def test_save_candidate_to_excel(self, mock_workbook):
        # Создаем фиктивные данные кандидата
        candidate_data = {'result': {'ID': 456, 'NAME': 'Петр', 'LAST_NAME': 'Петров', 'PHONE': [{'VALUE': '+79876543210'}], 'EMAIL': [{'VALUE': 'petrov@example.com'}], 'DATE_CREATE': '2023-10-02T13:00:00+0300'}}

        # Вызываем функцию
        result = save_candidate_to_excel(candidate_data)

        # Проверяем результат
        self.assertIn('Данные кандидата сохранены в candidates_', result)

    @patch('requests.post')
    def test_upload_file_to_lead_success(self, mock_requests_post):
        # Подготавливаем фиктивный ответ от API
        mock_response = MagicMock()
        mock_response.json.return_value = {'result': 42}
        mock_requests_post.return_value = mock_response

        # Вызываем функцию
        result = upload_file_to_lead('test_file.xlsx')

        # Проверяем результат
        self.assertEqual(result, 42)


    @patch('requests.post')
    def test_save_link_to_file_success(self, mock_requests_post):
        # Подготавливаем фиктивный ответ от API
        mock_response = MagicMock()
        mock_response.json.return_value = {}
        mock_requests_post.return_value = mock_response

        # Вызываем функцию
        save_link_to_file(42, '/path/to/file.xlsx', 123)

        # Проверяем, что запрос был выполнен с правильными параметрами
        expected_payload = {
            "id": 123,
            "fields": {
                42: {"value": "/path/to/file.xlsx"}
            }
        }
        mock_requests_post.assert_called_once_with(f'{BITRIX_WEBHOOK_URL}crm.lead.update.json', json=expected_payload)


    @patch('openpyxl.load_workbook')
    def test_read_from_excel_success(self, mock_load_workbook):
        # Подготавливаем фиктивный рабочий лист
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = [
            ('Title1', 'Name1', 'LastName1', 'Phone1', 'email1@example.com'),
            ('Title2', 'Name2', 'LastName2', 'Phone2', 'email2@example.com')
        ]
        mock_wb = MagicMock()
        type(mock_wb).active = mock_ws
        mock_load_workbook.return_value = mock_wb

        # Вызываем функцию
        result = read_from_excel('test.xlsx')

        # Проверяем результат
        expected_result = [
            {
                'TITLE': 'Title1',
                'NAME': 'Name1',
                'LAST_NAME': 'LastName1',
                'PHONE': 'Phone1',
                'EMAIL': 'email1@example.com'
            },
            {
                'TITLE': 'Title2',
                'NAME': 'Name2',
                'LAST_NAME': 'LastName2',
                'PHONE': 'Phone2',
                'EMAIL': 'email2@example.com'
            }
        ]
        self.assertEqual(result, expected_result)


    @patch('requests.post')
    def test_create_smart_process_success(self, mock_requests_post):
        # Подготавливаем фиктивный ответ от API
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_requests_post.return_value = mock_response

        # Вызываем функцию
        data = [
            {
                'TITLE': 'Title1',
                'NAME': 'Name1',
                'LAST_NAME': 'LastName1',
                'PHONE': 'Phone1',
                'EMAIL': 'email1@example.com'
            },
            {
                'TITLE': 'Title2',
                'NAME': 'Name2',
                'LAST_NAME': 'LastName2',
                'PHONE': 'Phone2',
                'EMAIL': 'email2@example.com'
            }
        ]
        create_smart_process(data)

        # Проверяем, что запрос был отправлен дважды
        self.assertEqual(mock_requests_post.call_count, 2)

    @patch('requests.post')
    def test_create_smart_process_failure(self, mock_requests_post):
        # Подготавливаем фиктивный ответ от API с ошибкой
        mock_response = MagicMock()
        mock_response.status_code = 400
        mock_response.text = 'Bad Request'
        mock_requests_post.return_value = mock_response

        # Вызываем функцию
        data = [
            {
                'TITLE': 'Title1',
                'NAME': 'Name1',
                'LAST_NAME': 'LastName1',
                'PHONE': 'Phone1',
                'EMAIL': 'email1@example.com'
            },
            {
                'TITLE': 'Title2',
                'NAME': 'Name2',
                'LAST_NAME': 'LastName2',
                'PHONE': 'Phone2',
                'EMAIL': 'email2@example.com'
            }
        ]
        create_smart_process(data)

        # Проверяем, что запрос был отправлен дважды
        self.assertEqual(mock_requests_post.call_count, 2)

if __name__ == '__main__':
    unittest.main()