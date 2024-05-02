import os
import psutil


def check_data(line_data_file, line_start_folder, line_finish_folder):

    for proc in psutil.process_iter():
        if proc.name() == 'WINWORD.EXE':
            return ['УПС!', 'Закройте все файлы Word!']
    path_file = line_data_file.text().strip()
    if not path_file:
        return ['УПС!', 'Путь к файлу выгрузки пуст']
    if os.path.exists(path_file) is False:
        return ['УПС!', 'Не удается найти указанный файл выгрузки']
    if os.path.isdir(path_file):
        return ['УПС!', 'Указанный путь к файлу выгрузки является директорией']
    if path_file.endswith('.txt'):
        pass
    else:
        return ['УПС!', 'Загружаемый файл выгрузки не формата ".txt"']
    start_path = line_start_folder.text().strip()
    if not start_path:
        return ['УПС!', 'Путь к папке с начальными документами пуст']
    if os.path.isfile(start_path):
        return ['УПС!', 'В пути к папке с начальными документами указан файл']
    if os.path.exists(start_path) is False:
        return ['УПС!', 'Не удается найти папку с начальными документами']
    finish_path = line_finish_folder.text().strip()
    if not finish_path:
        return ['УПС!', 'Путь к конечной папке пуст']
    if os.path.isfile(finish_path):
        return ['УПС!', 'В пути к конечной папке указан файл']
    if os.path.exists(finish_path) is False:
        return ['УПС!', 'Не удается найти указанную конечную директорию']
    if len(os.listdir(finish_path)) > 0:
        return ['УПС!', 'Конечная папка не пуста, очистите директорию или укажите другую']

    return {'path_file': path_file, 'start_path': start_path, 'finish_path': finish_path}

