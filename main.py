import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import scrolledtext
import configparser
import os
import sys
import requests
import time
import threading
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import shutil
import re
import json
import urllib.parse
import tempfile

# Объявляем глобальную переменную для отладочного логирования на уровне модуля
# Инициализируем ее значением по умолчанию. Реальное значение будет загружено из конфига.
global_debug_logging = False


# --- Конфигурация ---
CONFIG_FILE = "config.ini"
LOG_FILE_NAME = "debug_log.log"

# Значения конфигурации по умолчанию
DEFAULT_CONFIG = {
    'Settings': {
        'HttpRequestTimeoutSec': '15',
        'InstallerRoot': 'C:\\iiko_Distr', # Корневой каталог для ЛОКАЛЬНЫХ дистрибутивов
        'ConfigFileWaitTimeoutSec': '60',
        'ConfigFileCheckIntervalMs': '100',
        'DebugLogging': 'False' # Включить подробное логирование в консоль
    },
    # Определяем ПРИОРИТЕТ источников. Перечислять через запятую.
    # Скрипт будет проверять источники в указанном порядке.
    'SourcePriority': {
        'Order': 'smb, http, ftp'
    },
    # Настройки для SMB источника
    'SmbSource': {
        'Enabled': 'False', # Включить этот источник?
        'Path': '\\\\10.25.100.5\\sharedisk\\iikoBacks', # UNC-путь к корневой папке на SMB
        # Шаблоны имен архивов на SMB. {version} будет заменено на форматированную версию.
        # {vendor_subdir} будет заменено на "Syrve/" для Syrve и "" для iiko.
        # Важно: эти шаблоны относятся к именам ZIP-АРХИВОВ на SMB.
        'iikoRMS_ArchiveName': 'RMSOffice{version}.zip',
        'iikoChain_ArchiveName': 'ChainOffice{version}.zip',
        'SyrveRMS_ArchiveName': 'Syrve/RMSSOffice{version}.zip', # Указываем подпапку Syrve
        'SyrveChain_ArchiveName': 'Syrve/ChainSOffice{version}.zip' # Указываем подпапку Syrve
    },
    # Настройки для HTTP источника
    'HttpSource': {
        'Enabled': 'True', # Включить этот источник?
        'Url': 'https://f.serty.top/iikoBacks', # Базовый URL директории с архивами
        # Шаблоны имен архивов на HTTP. {version} будет заменено на форматированную версию.
        # {vendor_subdir} будет заменено на "Syrve/" для Syrve и "" для iiko.
        'iikoRMS_ArchiveName': 'RMSOffice{version}.zip',
        'iikoChain_ArchiveName': 'ChainOffice{version}.zip',
        'SyrveRMS_ArchiveName': 'Syrve/RMSSOffice{version}.zip', # Указываем подпапку Syrve
        'SyrveChain_ArchiveName': 'Syrve/ChainSOffice{version}.zip' # Указываем подпапку Syrve
    },
    # Настройки для FTP источника
    'FtpSource': {
        'Enabled': 'True', # Включить этот источник?
        'Host': 'ftp.serty.top',
        'Port': '21', # Стандартный порт FTP
        'Username': 'ftpuser',
        'Password': '11', # Внимание: хранение паролей в конфиге небезопасно!
        'Directory': '/iikoBacks', # Путь к директории с архивами на FTP сервере
        # Шаблоны имен архивов на FTP. {version} будет заменено на форматированную версию.
        # {vendor_subdir} будет заменено на "Syrve/" для Syrve и "" для iiko.
        'iikoRMS_ArchiveName': 'RMSOffice{version}.zip',
        'iikoChain_ArchiveName': 'ChainOffice{version}.zip',
        'SyrveRMS_ArchiveName': 'Syrve/RMSSOffice{version}.zip', # Указываем подпапку Syrve
        'SyrveChain_ArchiveName': 'Syrve/ChainSOffice{version}.zip' # Указываем подпапку Syrve
    },
    # Определяем ФОРМАТ имен ПАПОК для ЛОКАЛЬНОГО хранения дистрибутивов.
    # Это ИМЯ КАТАЛОГА, а не архива.
    'LocalInstallerNames': {
        'iikoRMS': 'RMSOffice',
        'iikoChain': 'ChainOffice',
        'SyrveRMS': 'RMSSOffice',
        'SyrveChain': 'ChainSOffice'
    }
}

# --- Вспомогательные функции ---

def load_config():
    """Загружает конфигурацию из config.ini или создает его со значениями по умолчанию."""
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        print(f"Создание файла конфигурации по умолчанию: {CONFIG_FILE}")
        try:
            # Создаем и заполняем конфиг в памяти из DEFAULT_CONFIG
            for section, values in DEFAULT_CONFIG.items():
                config.add_section(section)
                for key, value in values.items():
                    config.set(section, key, value)
            # Записываем в файл
            with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            print(f"Файл конфигурации '{CONFIG_FILE}' успешно создан.")
        except Exception as e:
            print(f"Ошибка при создании файла конфигурации по умолчанию '{CONFIG_FILE}': {e}", file=sys.stderr)
            # Если не удалось создать файл, работаем с конфигом в памяти из DEFAULT_CONFIG
            print("Используются значения конфигурации по умолчанию из памяти.")
            # Убедимся, что config объект содержит дефолты, даже если создание файла упало
            config = configparser.ConfigParser()
            for section, values in DEFAULT_CONFIG.items():
                 config.add_section(section)
                 for key, value in values.items():
                     config.set(section, key, value)

    else:
        try:
            config.read(CONFIG_FILE, encoding='utf-8')
            print(f"Файл конфигурации '{CONFIG_FILE}' успешно загружен.")
            # Проверяем наличие всех секций и ключей из DEFAULT_CONFIG и добавляем их, если отсутствуют
            # Это позволяет добавлять новые настройки в DEFAULT_CONFIG без необходимости удалять старый config.ini
            config_changed = False
            for section, values in DEFAULT_CONFIG.items():
                if not config.has_section(section):
                    config.add_section(section)
                    config_changed = True
                for key, value in values.items():
                    if not config.has_option(section, key):
                        config.set(section, key, value)
                        config_changed = True

            # Если были добавлены новые настройки, сохраняем обновленный файл
            if config_changed:
                 print(f"Обновление файла конфигурации '{CONFIG_FILE}' новыми параметрами.")
                 try:
                     with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                         config.write(configfile)
                     print(f"Файл конфигурации '{CONFIG_FILE}' успешно обновлен.")
                 except Exception as e:
                     print(f"Ошибка при обновлении файла конфигурации '{CONFIG_FILE}': {e}", file=sys.stderr)


        except Exception as e:
            print(f"Ошибка при загрузке файла конфигурации '{CONFIG_FILE}': {e}", file=sys.stderr)
            print("Используются значения конфигурации по умолчанию из памяти.")
            # Если загрузка упала, работаем с дефолтами в памяти
            config = configparser.ConfigParser()
            for section, values in DEFAULT_CONFIG.items():
                 config.add_section(section)
                 for key, value in values.items():
                     config.set(section, key, value)


    return config

def get_config_value(config, section, key, default=None, type_cast=str):
    """Безопасно получает значение конфигурации с приведением типа и значением по умолчанию."""
    try:
        if type_cast == bool:
            # getboolean возвращает bool
            return config.getboolean(section, key)
        elif type_cast == int:
            # getint возвращает int
            return config.getint(section, key)
        elif type_cast == float:
             # getfloat возвращает float
             return config.getfloat(section, key)
        else:
            # get возвращает str
            return config.get(section, key)
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        # Логируем предупреждение только если это не отладочный режим (чтобы избежать рекурсии при логировании)
        try:
            if not global_debug_logging:
                 print(f"Внимание: Не удалось получить значение конфигурации [{section}]{key}. Использование значения по умолчанию: {default}. Ошибка: {e}", file=sys.stderr)
        except NameError:
             # Если global_debug_logging еще не определена, просто печатаем
             print(f"Внимание: Не удалось получить значение конфигурации [{section}]{key}. Использование значения по умолчанию: {default}. Ошибка: {e}", file=sys.stderr)

        return default


def log_message(message, level="INFO"):
    """Выводит сообщение в консоль с временной меткой и уровнем."""
    # Объявляем, что используем глобальную переменную
    global global_debug_logging

    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    formatted_message = f"[{timestamp}] [{level}] {message}"

    # Проверяем уровень отладки перед выводом сообщения
    if level == "DEBUG" and not global_debug_logging:
        pass # Пропускаем отладочные сообщения, если не включено
    else:
        print(formatted_message, file=sys.stdout if level != "ERROR" else sys.stderr)
    
    if global_debug_logging:
        try:
            # Открываем файл в режиме добавления ('a'), создаем его, если не существует.
            # Используем кодировку UTF-8 для поддержки русских символов.
            with open(LOG_FILE_NAME, 'a', encoding='utf-8') as log_file:
                log_file.write(formatted_message + '\n')
        except Exception as e:
            # Если запись в лог-файл не удалась, выводим ошибку в консоль.
            # Важно: не используем log_message здесь, чтобы избежать рекурсии при ошибке записи лога.
            print(f"[{timestamp}] [ERROR] Ошибка записи в лог-файл '{LOG_FILE_NAME}': {e}", file=sys.stderr)


def parse_target_string(input_string):
    """Парсит входную строку (URL или IP:Port)."""
    log_message(f"DEBUG: Начат парсинг строки: '{input_string}'")
    if not input_string or not input_string.strip():
        log_message("Входная строка для парсинга пуста.", level="ERROR")
        return None

    url_or_ip = None
    port = 443 # Порт по умолчанию
    is_ip_address = False

    # Удаляем потенциальную схему и символы авторизации (@) для упрощения парсинга хоста/порта
    temp_input = re.sub(r"^(https?|ftp|ftps)://", "", input_string, flags=re.IGNORECASE)
    temp_input = re.sub(r"^[^@]+@", "", temp_input) # Удаляем user:pass@ если есть

    log_message(f"DEBUG: Строка после удаления схемы/авторизации: '{temp_input}'")

    # Разделяем хост и порт. Ищем последнее двоеточие, которое не является частью IPv6 адреса.
    # Для простоты пока ищем двоеточие, за которым следуют цифры до конца строки или слэша.
    port_match = re.search(r":(\d+)(?:/.*)?$", temp_input)
    if port_match:
        try:
            port = int(port_match.group(1))
            log_message(f"DEBUG: Порт '{port}' извлечен из исходной строки.")
            # Удаляем часть с портом из строки для парсинга хоста
            temp_input_for_host = temp_input[:port_match.start()]
        except ValueError:
            log_message(f"DEBUG: Не удалось распарсить порт из '{port_match.group(1)}', используется порт по умолчанию 443.", level="WARNING")
            temp_input_for_host = temp_input # Используем исходную строку без удаления порта
            port = 443
    else:
         log_message("DEBUG: Явный порт не найден, используется порт по умолчанию 443.")
         temp_input_for_host = temp_input # Нет порта для удаления


    # Удаляем путь, если он есть (часть после первого слэша) из строки для хоста
    host_part = temp_input_for_host.split('/', 1)[0]
    url_or_ip = host_part

    # Простая проверка, является ли это IPv4 адресом
    if re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$", url_or_ip):
         # Дополнительная валидация октетов
         octets = url_or_ip.split('.')
         if len(octets) == 4 and all(0 <= int(octet) <= 255 for octet in octets):
             is_ip_address = True
             log_message("DEBUG: Хост определен как IP-адрес IPv4.")
         else:
             log_message(f"DEBUG: Хост '{url_or_ip}' выглядит как IP, но октеты неверны или их количество не 4.", level="WARNING")
    # Можно добавить проверку на IPv6, но для простоты пока ограничимся IPv4

    log_message(f"DEBUG: Парсинг завершен: Хост/IP: '{url_or_ip}', Порт: {port}, IsIpAddress: {is_ip_address}")

    return {
        'UrlOrIp': url_or_ip,
        'Port': port,
        'IsIpAddress': is_ip_address
    }


def format_version(version_string):
    """Форматирует строку версии, возвращая только первые цифры из трёх первых частей."""
    log_message(f"DEBUG: Форматирование версии: '{version_string}'")
    if not version_string:
        log_message("Входная строка версии пуста.", level="WARNING")
        return ""

    parts = version_string.split('.')
    digits = []

    for part in parts:
        # Берем только начальные цифры из каждой части
        match = re.match(r"^\d+", part)
        if match:
            # Добавляем только ПЕРВУЮ цифру из найденного числового сегмента
            digits.append(match.group(0)[0])
        # Если мы уже получили цифры из 3 частей, останавливаемся
        if len(digits) == 3:
            break

    # Объединяем собранные цифры (максимум 3)
    result = "".join(digits)

    # Если после попытки извлечения цифр результат пустой, но исходная строка была не пуста
    if not result and version_string:
        log_message(f"Не удалось извлечь первые цифры из первых трех частей версии '{version_string}'. Возвращаем исходную строку.", level="WARNING")
        return version_string # Возвращаем исходную строку

    log_message(f"DEBUG: Форматированная версия: '{result}'")
    return result


def determine_app_type(input_string, edition):
    """Определяет тип приложения и производителя на основе входной строки и edition."""
    log_message(f"DEBUG: Определение типа приложения для строки '{input_string}' и edition '{edition}'")
    vendor = "iiko" # По умолчанию
    # Проверяем входную строку на наличие "syrve" (без учета регистра)
    if "syrve" in input_string.lower():
        vendor = "Syrve"

    app_type = None
    # Определяем тип на основе edition
    if edition: # Проверяем, что edition не None или пустая строка
        if edition.lower() == "default":
            app_type = f"{vendor}RMS"
        elif edition.lower() == "chain":
            app_type = f"{vendor}Chain"
        else:
            # Не удалось определить автоматически, требуется взаимодействие с пользователем.
            # В GUI это будет обработано вызовом диалога.
            log_message(f"Не удалось автоматически определить тип RMS/Chain по edition ('{edition}') для производителя '{vendor}'. Требуется выбор пользователя.", level="WARNING")
            return None # Возвращаем None, чтобы сигнализировать о необходимости выбора
    else:
        log_message(f"Значение 'edition' из ответа сервера пустое или отсутствует. Не удалось автоматически определить тип приложения.", level="WARNING")
        return None # Требуется выбор пользователя, так как edition не определен

    log_message(f"DEBUG: Определен тип приложения: '{app_type}' (Производитель: '{vendor}')")
    return {
        'AppType': app_type,
        'Vendor': vendor
    }

def get_expected_installer_name(config, app_type, version_formatted):
    """Формирует ожидаемое имя ЛОКАЛЬНОГО каталога дистрибутива на основе конфига."""
    log_message(f"DEBUG: Формирование имени ЛОКАЛЬНОГО каталога дистрибутива для типа '{app_type}' и версии '{version_formatted}'")

    # Читаем базовое имя из конфига LocalInstallerNames
    base_name = get_config_value(config, 'LocalInstallerNames', app_type, default=None, type_cast=str)

    if base_name is None:
        log_message(f"Ошибка: Не найдено базовое имя ЛОКАЛЬНОГО каталога для типа приложения '{app_type}' в разделе LocalInstallerNames конфига.", level="ERROR")
        return None

    result_name = f"{base_name}{version_formatted}"
    log_message(f"DEBUG: Ожидаемое имя ЛОКАЛЬНОГО каталога дистрибутива: '{result_name}'")
    return result_name

def sanitize_for_path(input_string):
    """Очищает строку для использования в пути к файлу, удаляя или заменяя недопустимые символы."""
    log_message(f"DEBUG: Санитизация строки для пути: '{input_string}'")
    # Заменяем двоеточие на тире (часто используется в URL:порт)
    sanitized = input_string.replace(':', '-')
    # Удаляем недопустимые символы Windows
    sanitized = re.sub(r'[<>"/\\|?*]', '_', sanitized)
    # Удаляем начальные/конечные пробелы (ограничение Windows)
    sanitized = re.sub(r'^[.\s]+|[.\s]+$', '', sanitized)
    # Схлопываем последовательности точек, подчеркиваний и пробелов в одно подчеркивание
    sanitized = re.sub(r'[._\s]+', '_', sanitized)
     # Удаляем конечные точки (строгое правило Windows)
    sanitized = re.sub(r'\.+$', '', sanitized)

    # Убеждаемся, что строка не пустая после санитизации
    if not sanitized:
        sanitized = "default_name"
        log_message(f"Внимание: Строка стала пустой после санитизации. Используется имя по умолчанию: '{sanitized}'", level="WARNING")
    log_message(f"DEBUG: Санитизированная строка: '{sanitized}'")
    return sanitized

import ctypes
import ctypes.wintypes

def get_file_company_name(filepath):
    """Получает CompanyName из свойств файла через WinAPI (только для Windows)."""
    # Проверяем, что мы на Windows
    if os.name != 'nt':
        log_message(f"DEBUG: get_file_company_name через WinAPI доступна только на Windows. Текущая ОС: {os.name}", level="DEBUG")
        return None # Возвращаем None, если не на Windows

    log_message(f"DEBUG: Попытка чтения CompanyName из файла через WinAPI: '{filepath}'")

    if not os.path.isfile(filepath):
        log_message(f"Ошибка: Файл не найден для чтения метаданных: '{filepath}'", level="ERROR")
        return None

    try:
        # Получаем размер блока версии
        size = ctypes.windll.version.GetFileVersionInfoSizeW(filepath, None)
        if size == 0:
            log_message(f"DEBUG: Информация о версии отсутствует в файле '{filepath}'.", level="DEBUG")
            return None

        # Читаем блок версии
        res = ctypes.create_string_buffer(size)
        # Внимание: GetFileVersionInfoW может вернуть 0 при ошибке, но ctypes.windll не генерирует исключение по умолчанию.
        # Для более надежной обработки ошибок WinAPI можно использовать check_last_error=True
        if ctypes.windll.version.GetFileVersionInfoW(filepath, 0, size, res) == 0:
             last_error = ctypes.GetLastError()
             log_message(f"Ошибка WinAPI GetFileVersionInfoW для файла '{filepath}': Код ошибки {last_error}", level="ERROR")
             return None


        # Получаем список переводов (язык + кодировка)
        lplpBuffer = ctypes.c_void_p()
        puLen = ctypes.wintypes.UINT()
        # Внимание: VerQueryValueW может вернуть 0 при ошибке
        if ctypes.windll.version.VerQueryValueW(res, r'\\VarFileInfo\\Translation', ctypes.byref(lplpBuffer), ctypes.byref(puLen)) == 0:
             last_error = ctypes.GetLastError()
             log_message(f"Ошибка WinAPI VerQueryValueW (Translation) для файла '{filepath}': Код ошибки {last_error}", level="ERROR")
             # Попробуем использовать стандартный перевод 040904b0 (английский) если список переводов недоступен
             log_message("DEBUG: Не удалось получить список переводов. Попытка использовать стандартный перевод 040904b0.", level="DEBUG")
             lang_codepage = '040904b0' # Английский (США) + Unicode
        else:
             # Извлекаем первый перевод
             translation = ctypes.cast(lplpBuffer, ctypes.POINTER(ctypes.c_ushort * 2)).contents
             lang_codepage = f'{translation[0]:04x}{translation[1]:04x}'
             log_message(f"DEBUG: Найден перевод: {lang_codepage}", level="DEBUG")


        # Формируем путь к нужному значению
        sub_block = f'\\StringFileInfo\\{lang_codepage}\\CompanyName'

        # Получаем строку производителя
        lplpBuffer = ctypes.c_wchar_p() # Используем c_wchar_p для строк WinAPI (UTF-16)
        puLen = ctypes.wintypes.UINT()
        # Внимание: VerQueryValueW может вернуть 0 при ошибке
        if ctypes.windll.version.VerQueryValueW(res, sub_block, ctypes.byref(lplpBuffer), ctypes.byref(puLen)) == 0:
             last_error = ctypes.GetLastError()
             log_message(f"DEBUG: Ошибка WinAPI VerQueryValueW (CompanyName) для файла '{filepath}': Код ошибки {last_error}", level="DEBUG")
             log_message(f"CompanyName не найдено в метаданных файла '{filepath}' (или ошибка чтения).", level="WARNING")
             return None


        company_name = lplpBuffer.value
        log_message(f"DEBUG: Извлечено CompanyName: '{company_name}'", level="DEBUG")
        return company_name.strip() if company_name else None

    except Exception as e:
        log_message(f"Ошибка при извлечении CompanyName через WinAPI для файла '{filepath}': {e}", level="ERROR")
        return None


def _download_from_http(config, app_type, version_formatted, expected_installer_name, temp_archive_path, update_status, update_progress_callback, base_progress, progress_range):
    """Скачивает архив дистрибутива по HTTP."""
    log_message(f"DEBUG: Попытка скачивания с HTTP.")
    http_enabled = get_config_value(config, 'HttpSource', 'Enabled', default=False, type_cast=bool)
    if not http_enabled:
        log_message("DEBUG: HTTP источник отключен в конфиге.", level="DEBUG")
        return False # Источник отключен

    http_url_base = get_config_value(config, 'HttpSource', 'Url', default=None, type_cast=str)
    if not http_url_base:
        log_message("Ошибка: Не указан URL для HTTP источника в конфиге.", level="ERROR")
        return False # Не настроен

    # Определяем имя архива для HTTP источника
    archive_name_template = get_config_value(config, 'HttpSource', f'{app_type}_ArchiveName', default=None, type_cast=str)
    if not archive_name_template:
         log_message(f"Ошибка: Не указан шаблон имени архива для типа '{app_type}' в разделе HttpSource конфига.", level="ERROR")
         return False # Не настроен

    # Заменяем {version} на форматированную версию
    archive_name = archive_name_template.replace('{version}', version_formatted)

    # Соединяем базовый URL и имя архива. Используем urljoin для корректной обработки путей.
    http_full_url = urllib.parse.urljoin(http_url_base.rstrip('/') + '/', archive_name)

    update_status(f"Скачивание с HTTP: {os.path.basename(http_full_url)}...")
    log_message(f"Попытка скачивания с HTTP: '{http_full_url}' в '{temp_archive_path}'.")

    try:
        # Используем stream=True для скачивания больших файлов по частям
        response = requests.get(http_full_url, stream=True, timeout=get_config_value(config, 'Settings', 'HttpRequestTimeoutSec', default=15, type_cast=int))
        response.raise_for_status() # Генерирует исключение для плохих кодов статуса (4xx или 5xx)

        total_size = int(response.headers.get('content-length', 0))
        downloaded_size = 0
        buffer_size = 8192 # Размер буфера для чтения/записи

        with open(temp_archive_path, 'wb') as f_dst:
            for chunk in response.iter_content(chunk_size=buffer_size):
                if chunk: # filter out keep-alive new chunks
                    f_dst.write(chunk)
                    downloaded_size += len(chunk)
                    # Обновление прогресса
                    if total_size > 0:
                         progress_value = base_progress + (downloaded_size / total_size) * progress_range
                         update_progress_callback(progress_value)

        log_message("Скачивание HTTP завершено.")
        update_status("Скачивание HTTP завершено.")
        update_progress_callback(base_progress + progress_range) # Убедимся, что прогресс достигает конца этапа
        return True # Успех

    except requests.exceptions.RequestException as e:
        log_message(f"Ошибка HTTP скачивания с '{http_full_url}': {e}", level="ERROR")
        update_status(f"Ошибка HTTP скачивания: {e}", level="ERROR")
        return False # Ошибка скачивания
    except Exception as e:
        log_message(f"Неизвестная ошибка при скачивании с HTTP '{http_full_url}': {e}", level="ERROR")
        update_status(f"Неизвестная ошибка HTTP скачивания: {e}", level="ERROR")
        return False # Неизвестная ошибка


def _download_from_ftp(config, app_type, version_formatted, expected_installer_name, temp_archive_path, update_status, update_progress_callback, base_progress, progress_range):
    """Скачивает архив дистрибутива по FTP."""
    log_message(f"DEBUG: Попытка скачивания с FTP.")
    ftp_enabled = get_config_value(config, 'FtpSource', 'Enabled', default=False, type_cast=bool)
    if not ftp_enabled:
        log_message("DEBUG: FTP источник отключен в конфиге.", level="DEBUG")
        return False # Источник отключен

    ftp_host = get_config_value(config, 'FtpSource', 'Host', default=None, type_cast=str)
    ftp_port = get_config_value(config, 'FtpSource', 'Port', default=21, type_cast=int)
    ftp_username = get_config_value(config, 'FtpSource', 'Username', default='anonymous', type_cast=str)
    ftp_password = get_config_value(config, 'FtpSource', 'Password', default='', type_cast=str)
    ftp_directory = get_config_value(config, 'FtpSource', 'Directory', default=None, type_cast=str)

    if not ftp_host or not ftp_directory:
        log_message("Ошибка: Не указаны Host или Directory для FTP источника в конфиге.", level="ERROR")
        return False # Не настроен

    # Определяем имя архива для FTP источника
    archive_name_template = get_config_value(config, 'FtpSource', f'{app_type}_ArchiveName', default=None, type_cast=str)
    if not archive_name_template:
         log_message(f"Ошибка: Не указан шаблон имени архива для типа '{app_type}' в разделе FtpSource конфига.", level="ERROR")
         return False # Не настроен

    # Заменяем {version} на форматированную версию
    archive_name = archive_name_template.replace('{version}', version_formatted)

    # Соединяем директорию и имя архива на FTP. ftplib требует отдельно директорию и имя файла.
    ftp_full_path = f"{ftp_directory.rstrip('/')}/{archive_name}"


    update_status(f"Скачивание с FTP: {archive_name}...")
    log_message(f"Попытка скачивания с FTP: '{ftp_host}:{ftp_port}{ftp_full_path}' в '{temp_archive_path}'.")

    from ftplib import FTP
    try:
        with FTP() as ftp:
            ftp.connect(ftp_host, ftp_port)
            ftp.login(ftp_username, ftp_password)
            log_message(f"DEBUG: FTP логин успешен. Текущая директория: {ftp.pwd()}", level="DEBUG")

            # Переходим в нужную директорию
            # ftplib.cwd может принимать относительные пути, включая подпапки типа "Syrve/"
            ftp.cwd(ftp_directory)
            log_message(f"DEBUG: FTP смена директории на '{ftp_directory}' успешна. Текущая директория: {ftp.pwd()}", level="DEBUG")

            # Если archive_name_template включал подпапку (например "Syrve/"),
            # то archive_name будет "Syrve/RMSSOffice887.zip".
            # Нам нужно сменить директорию еще раз для Syrve
            archive_dir = os.path.dirname(archive_name)
            archive_file = os.path.basename(archive_name)

            if archive_dir and archive_dir != '.': # Если есть подпапка в имени архива
                 try:
                      ftp.cwd(archive_dir)
                      log_message(f"DEBUG: FTP смена директории на подпапку '{archive_dir}' успешна. Текущая директория: {ftp.pwd()}", level="DEBUG")
                 except Exception as e:
                      log_message(f"Ошибка FTP: Не удалось сменить директорию на '{archive_dir}': {e}", level="ERROR")
                      update_status(f"Ошибка FTP: Не найдена подпапка '{archive_dir}'.", level="ERROR")
                      return False # Ошибка

            # Получаем размер файла для прогресса
            try:
                total_size = ftp.size(archive_file)
                log_message(f"DEBUG: Размер архива на FTP: {total_size} байт.", level="DEBUG")
            except Exception as e:
                 log_message(f"Ошибка FTP: Не удалось получить размер файла '{archive_file}': {e}", level="WARNING")
                 total_size = 0


            downloaded_size = 0
            buffer_size = 8192

            # Callback функция для передачи прогресса в retrbinary
            def handle_ftp_progress(chunk):
                 nonlocal downloaded_size
                 downloaded_size += len(chunk)
                 if total_size > 0:
                      progress_value = base_progress + (downloaded_size / total_size) * progress_range
                      update_progress_callback(progress_value)
                 f_dst.write(chunk)


            with open(temp_archive_path, 'wb') as f_dst:
                 # ftp.retrbinary('RETR filename', callback, blocksize)
                 ftp.retrbinary(f'RETR {archive_file}', handle_ftp_progress, buffer_size)

            log_message("Скачивание FTP завершено.")
            update_status("Скачивание FTP завершено.")
            update_progress_callback(base_progress + progress_range) # Убедимся, что прогресс достигает конца этапа
            return True # Успех

    except Exception as e:
        log_message(f"Ошибка FTP скачивания с '{ftp_host}:{ftp_port}{ftp_full_path}': {e}", level="ERROR")
        update_status(f"Ошибка FTP скачивания: {e}", level="ERROR")
        return False # Ошибка скачивания


def _download_from_smb(config, app_type, version_formatted, expected_installer_name, temp_archive_path, update_status, update_progress_callback, base_progress, progress_range):
    """Скачивает архив дистрибутива с SMB ресурса (копированием)."""
    log_message(f"DEBUG: Попытка скачивания с SMB.")
    smb_enabled = get_config_value(config, 'SmbSource', 'Enabled', default=False, type_cast=bool)
    if not smb_enabled:
        log_message("DEBUG: SMB источник отключен в конфиге.", level="DEBUG")
        return False # Источник отключен

    smb_path_base = get_config_value(config, 'SmbSource', 'Path', default=None, type_cast=str)
    if not smb_path_base:
        log_message("Ошибка: Не указан Path для SMB источника в конфиге.", level="ERROR")
        return False # Не настроен

    # Определяем имя архива для SMB источника
    archive_name_template = get_config_value(config, 'SmbSource', f'{app_type}_ArchiveName', default=None, type_cast=str)
    if not archive_name_template:
         log_message(f"Ошибка: Не указан шаблон имени архива для типа '{app_type}' в разделе SmbSource конфига.", level="ERROR")
         return False # Не настроен

    # Заменяем {version} на форматированную версию
    archive_name = archive_name_template.replace('{version}', version_formatted)

    # Соединяем базовый путь и имя архива.
    smb_full_path = os.path.join(smb_path_base, archive_name)
    # При необходимости, os.path.join может некорректно работать с UNC путями и подпапками типа "Syrve/".
    # Более надежно использовать Path.Combine в .NET или вручную склеивать, но для простых случаев os.path.join может сработать.
    # Если шаблон включает подпапку (e.g., "Syrve/"), os.path.join(r"\\server\share", "Syrve/file.zip") может дать r"\\server\share\Syrve/file.zip"
    # Лучше вручную склеить, учитывая слеши
    smb_full_path = f"{smb_path_base.rstrip('/\\')}{os.sep}{archive_name.replace('/', os.sep).replace('\\', os.sep)}"


    update_status(f"Скачивание с SMB: {os.path.basename(smb_full_path)}...")
    log_message(f"Попытка скачивания с SMB: '{smb_full_path}' в '{temp_archive_path}'.")

    try:
        # Проверим существование исходного файла на SMB
        if not os.path.exists(smb_full_path):
             log_message(f"Ошибка: Исходный архив не найден на SMB: '{smb_full_path}'.", level="ERROR")
             update_status("Ошибка: Исходный архив не найден на SMB.", level="ERROR")
             return False

        # Получаем размер файла для прогресса
        try:
            total_size = os.path.getsize(smb_full_path)
            log_message(f"DEBUG: Размер архива на SMB: {total_size} байт.", level="DEBUG")
        except Exception as e:
             log_message(f"Ошибка при получении размера файла '{smb_full_path}': {e}", level="WARNING")
             total_size = 0

        copied_size = 0
        buffer_size = 1024 * 1024 # 1 MB buffer

        # Копирование файла по частям для индикации прогресса
        with open(smb_full_path, 'rb') as f_src, open(temp_archive_path, 'wb') as f_dst:
            while True:
                buffer = f_src.read(buffer_size)
                if not buffer:
                    break
                f_dst.write(buffer)
                copied_size += len(buffer)
                # Обновление прогресса
                if total_size > 0:
                    progress_value = base_progress + (copied_size / total_size) * progress_range
                    update_progress_callback(progress_value)

        log_message("Копирование SMB завершено.")
        update_status("Копирование SMB завершено.")
        update_progress_callback(base_progress + progress_range) # Убедимся, что прогресс достигает конца этапа
        return True # Успех

    except FileNotFoundError:
        log_message(f"Ошибка FileNotFoundError при скачивании с SMB: '{smb_full_path}'.", level="ERROR")
        update_status("Ошибка SMB скачивания: Файл не найден.", level="ERROR")
        return False # Ошибка
    except PermissionError:
        log_message(f"Ошибка доступа PermissionError при скачивании с SMB: '{smb_full_path}'. Проверьте права.", level="ERROR")
        update_status("Ошибка SMB скачивания: Нет прав доступа.", level="ERROR")
        return False # Ошибка
    except Exception as e:
        log_message(f"Неизвестная ошибка при скачивании с SMB '{smb_full_path}': {e}", level="ERROR")
        update_status(f"Неизвестная ошибка SMB скачивания: {e}", level="ERROR")
        return False # Неизвестная ошибка


# ... (вспомогательные функции скачивания _download_from_smb, _download_from_http, _download_from_ftp) ...

# --- Основная функция поиска/скачивания ---

def find_or_download_installer(config, app_type, version_formatted, vendor, update_status, update_progress_callback):
    """
    Находит дистрибутив локально или скачивает/распаковывает его с настроенных источников
    в порядке приоритета.
    """
    log_message(f"DEBUG: Начат поиск или скачивание дистрибутива для типа '{app_type}' версии '{version_formatted}' (производитель '{vendor}')")

    installer_root = get_config_value(config, 'Settings', 'InstallerRoot', default='D:\\Backs')
    # Определяем ожидаемое имя локальной папки
    expected_local_dir_name = get_expected_installer_name(config, app_type, version_formatted)
    if expected_local_dir_name is None:
         log_message("Ошибка определения имени локальной папки дистрибутива.", level="ERROR")
         update_status("Ошибка определения имени локальной папки.", level="ERROR")
         # update_progress_callback(0) # Прогресс уже 0
         return None

    local_installer_path = os.path.join(installer_root, expected_local_dir_name)
    backoffice_exe_direct_path = os.path.join(local_installer_path, "BackOffice.exe")

    update_status(f"Проверка локального дистрибутива: {expected_local_dir_name}...")

    # 1. Проверяем локально
    if os.path.exists(backoffice_exe_direct_path):
        log_message(f"Найден локальный дистрибутив: {local_installer_path}")
        company_name = get_file_company_name(backoffice_exe_direct_path) # Проверяем производителя
        # Если get_file_company_name вернул None (ctypes не работает или нет данных),
        # мы доверяем имени папки и считаем производителя совпадающим.
        if company_name is None or vendor.lower() in company_name.lower():
            update_status("Локальный дистрибутив найден и производитель совпадает (или не определен).")
            log_message("Производитель совпадает (или не определен). Используем локальный дистрибутив.")
            # Прогресс 100%, т.к. ничего скачивать не нужно
            update_progress_callback(100)
            return local_installer_path
        else:
            update_status("Локальный дистрибутив найден, но производитель не совпадает.", level="WARNING")
            log_message(f"Производитель локального дистрибутива ('{company_name}') не совпадает с ожидаемым ('{vendor}'). Будет предпринята попытка скачать.", level="WARNING")
            # Если производитель не совпадает, не используем локальную версию и пытаемся скачать.
            # Очищаем локальную папку, чтобы распаковать туда новую версию.
            if os.path.exists(local_installer_path):
                 log_message(f"Удаление локальной папки с несовпадающим производителем: '{local_installer_path}'")
                 shutil.rmtree(local_installer_path, ignore_errors=True)


    # 2. Если локальная проверка не удалась, начинаем процесс скачивания/распаковки/проверки
    update_status(f"Локальный дистрибутив не найден или не подходит. Попытка скачать с удаленных источников...")
    log_message(f"Локальный дистрибутив '{backoffice_exe_direct_path}' не найден или не прошел проверку.")

    # Инициализируем пути временных файлов/папок перед try блоком для доступа в except
    temp_archive_path = os.path.join(tempfile.gettempdir(), f"{expected_local_dir_name}.zip")
    temp_archive_path_exists = False # Флаг, был ли временный файл создан (хотя бы частично)
    temp_extract_path = os.path.join(local_installer_path, "temp_extract_folder")


    # --- ГЛАВНЫЙ TRY БЛОК для скачивания, распаковки и подготовки ---
    try:
        # Очищаем локальную папку на всякий случай, если она осталась в некорректном состоянии
        if os.path.exists(local_installer_path):
            log_message(f"DEBUG: Очистка существующей локальной папки '{local_installer_path}' перед скачиванием/распаковкой.", level="DEBUG")
            shutil.rmtree(local_installer_path, ignore_errors=True)

        # Создаем корневую папку дистрибутивов и папку для этого дистрибутива
        os.makedirs(installer_root, exist_ok=True)
        os.makedirs(local_installer_path, exist_ok=True)
        log_message(f"DEBUG: Создана локальная папка дистрибутива: '{local_installer_path}'.", level="DEBUG")

        # --- ДОБАВЛЕНО: Создаем родительскую директорию для временного архива ---
        # Это решает проблему FileNotFoundError, если путь к временному файлу включает несуществующие подпапки
        temp_archive_dir = os.path.dirname(temp_archive_path)
        if not os.path.exists(temp_archive_dir):
            os.makedirs(temp_archive_dir, exist_ok=True)
            log_message(f"DEBUG: Создана родительская директория для временного архива: '{temp_archive_dir}'.", level="DEBUG")
        # --- КОНЕЦ ДОБАВЛЕНОГО БЛОКА ---


        # 2.1. Скачиваем с удаленных источников по приоритету
        source_order_str = get_config_value(config, 'SourcePriority', 'Order', default='smb, http, ftp', type_cast=str)
        source_order = [s.strip().lower() for s in source_order_str.split(',') if s.strip()]

        download_success = False
        # Очищаем временный файл, если он вдруг остался от предыдущих попыток
        if os.path.exists(temp_archive_path):
             try:
                 os.remove(temp_archive_path)
                 log_message(f"DEBUG: Удален старый временный архив '{temp_archive_path}'.", level="DEBUG")
             except Exception as e:
                 log_message(f"Внимание: Не удалось удалить старый временный архив '{temp_archive_path}': {e}", level="WARNING")


        # Прогресс: 55% - 85% на скачивание и распаковку (30%)
        # Распределим 15% на скачивание и 15% на распаковку
        download_progress_base = 55
        download_progress_range = 15
        extract_progress_base = 70
        extract_progress_range = 15


        for source_type in source_order:
            log_message(f"DEBUG: Попытка скачивания с источника '{source_type}'...", level="DEBUG")
            if source_type == 'smb':
                update_status(f"Попытка скачивания с SMB...")
                if _download_from_smb(config, app_type, version_formatted, expected_local_dir_name, temp_archive_path, update_status, update_progress_callback, download_progress_base, download_progress_range):
                     download_success = True
                     temp_archive_path_exists = True # Файл создан (полностью или частично)
                     break # Скачивание успешно, переходим к распаковке
            elif source_type == 'http':
                update_status(f"Попытка скачивания с HTTP...")
                if _download_from_http(config, app_type, version_formatted, expected_local_dir_name, temp_archive_path, update_status, update_progress_callback, download_progress_base, download_progress_range):
                     download_success = True
                     temp_archive_path_exists = True # Файл создан (полностью или частично)
                     break # Скачивание успешно, переходим к распаковке
            elif source_type == 'ftp':
                update_status(f"Попытка скачивания с FTP...")
                if _download_from_ftp(config, app_type, version_formatted, expected_local_dir_name, temp_archive_path, update_status, update_progress_callback, download_progress_base, download_progress_range):
                     download_success = True
                     temp_archive_path_exists = True # Файл создан (полностью или частично)
                     break # Скачивание успешно, переходим к распаковке
            else:
                log_message(f"Внимание: Неизвестный источник в приоритете: '{source_type}'. Пропускаем.", level="WARNING")
                update_status(f"Неизвестный источник: '{source_type}'.", level="WARNING")

            log_message(f"DEBUG: Скачивание с источника '{source_type}' не удалось.", level="DEBUG")


        if not download_success:
            # Если цикл завершился без успешного скачивания
            raise RuntimeError("Не удалось скачать дистрибутив ни с одного доступного источника.")


        # 2.2. Распаковываем скачанный архив
        update_status(f"Распаковка архива '{os.path.basename(temp_archive_path)}'...")
        log_message(f"Распаковка архива '{temp_archive_path}' во временную папку '{temp_extract_path}'.")

        # Создаем временную папку для распаковки
        os.makedirs(temp_extract_path, exist_ok=True)

        try:
            with zipfile.ZipFile(temp_archive_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                total_files = len(file_list)
                extracted_count = 0

                if total_files == 0:
                     log_message("DEBUG: Архив пуст. Распаковка не требуется.", level="WARNING")

                for file_info in file_list:
                    # Распаковываем каждый файл во временную папку
                    zip_ref.extract(file_info, temp_extract_path)
                    extracted_count += 1
                    # Обновление прогресса
                    if total_files > 0: # Избегаем деления на ноль
                        # Прогресс внутри распаковки: (extracted_count / total_files) * range
                        progress_value = extract_progress_base + (extracted_count / total_files) * extract_progress_range
                        update_progress_callback(progress_value)
                    # update_status(f"Распаковка файла {extracted_count}/{total_files}...") # Слишком частые обновления

            log_message("Распаковка завершена.")
            update_status("Архив успешно распакован во временную папку.")
            update_progress_callback(extract_progress_base + extract_progress_range) # Убедимся, что прогресс достигает конца этапа распаковки


        except zipfile.BadZipFile:
            # Специфическая ошибка парсинга ZIP - перебрасываем с более понятным сообщением
            raise zipfile.BadZipFile(f"Архив '{os.path.basename(temp_archive_path)}' поврежден или не является ZIP-файлом.")
        # Любые другие ошибки при распаковке будут пойманы внешним except Exception


        # 2.3. Проверка и перемещение распакованного дистрибутива
        update_status("Проверка и перемещение распакованного дистрибутива...")
        log_message(f"Поиск BackOffice.exe во временной папке распаковки: '{temp_extract_path}'")

        found_backoffice_exe = None
        actual_content_root = None # Папка внутри temp_extract_path, которая является корнем дистрибутива

        # Ищем BackOffice.exe рекурсивно внутри временной папки распаковки
        for dirpath, dirnames, filenames in os.walk(temp_extract_path):
             if "BackOffice.exe" in filenames:
                 found_backoffice_exe = os.path.join(dirpath, "BackOffice.exe")
                 actual_content_root = dirpath # Папка, в которой найден BackOffice.exe
                 log_message(f"BackOffice.exe найден по пути: '{found_backoffice_exe}'. Корень содержимого: '{actual_content_root}'")
                 break # Нашли, останавливаем поиск

        if not found_backoffice_exe or not actual_content_root:
             # Если BackOffice.exe не найден после распаковки
             raise FileNotFoundError(f"Файл BackOffice.exe не найден в распакованном содержимом архива '{os.path.basename(temp_archive_path)}'.")


        # Теперь перемещаем содержимое actual_content_root в local_installer_path
        log_message(f"Перемещение содержимого из '{actual_content_root}' в '{local_installer_path}'")
        update_status("Перемещение содержимого дистрибутива...")
        try:
            # Перемещаем все элементы из найденной корневой папки во временной папке
            items_to_move = [os.path.join(actual_content_root, name) for name in os.listdir(actual_content_root)]
            for item_src in items_to_move:
                item_dest = os.path.join(local_installer_path, os.path.basename(item_src))
                # Используем shutil.move, который работает как переименование внутри одной ФС
                shutil.move(item_src, item_dest)

            log_message("DEBUG: Содержимое успешно перемещено.")
            update_status("Содержимое дистрибутива перемещено.")

        except Exception as e:
            # Ошибка при перемещении файлов
            raise RuntimeError(f"Ошибка при перемещении содержимого дистрибутива: {e}")


        # 2.4. Финальная проверка и верификация производителя BackOffice.exe после перемещения
        backoffice_exe_final_path = os.path.join(local_installer_path, "BackOffice.exe")
        if not os.path.exists(backoffice_exe_final_path):
             # Это не должно произойти после успешного перемещения, но на всякий случай
             raise FileNotFoundError(f"Файл BackOffice.exe не найден по ожидаемому конечному пути '{backoffice_exe_final_path}' после перемещения.")

        company_name = get_file_company_name(backoffice_exe_final_path)
        # Если get_file_company_name вернул None (ctypes не работает или нет данных),
        # мы доверяем производителю, определенному ранее (на основе имени архива/URL)
        if company_name is not None and vendor.lower() not in company_name.lower():
            # Производитель не совпадает
            raise ValueError(f"Производитель распакованного дистрибутива ('{company_name}') не совпадает с ожидаемым ('{vendor}'). Дистрибутив, возможно, некорректен.")


        # --- УСПЕХ: Удаляем временные папки и возвращаем путь ---
        log_message("Производитель распакованного дистрибутива совпадает (или не определен). Дистрибутив готов к использованию.")
        update_status("Дистрибутив успешно подготовлен.")

        if os.path.exists(temp_extract_path): # Удаляем временную папку распаковки
             try:
                 shutil.rmtree(temp_extract_path)
                 log_message("Временная папка распаковки удалена после успеха.")
             except Exception as e:
                 log_message(f"Ошибка при удалении временной папки распаковки '{temp_extract_path}' после успеха: {e}", level="WARNING")

        if temp_archive_path_exists and os.path.exists(temp_archive_path): # Удаляем временный архив
             try:
                 os.remove(temp_archive_path)
                 log_message("Временный архив успешно удален после успешной проверки.")
             except Exception as e:
                 log_message(f"Ошибка при удалении временного архива '{temp_archive_path}' после успеха: {e}", level="WARNING")

        update_progress_callback(100) # Убедимся, что прогресс 100%
        return local_installer_path # Возвращаем путь к готовому дистрибутиву


    # --- ГЛАВНЫЙ EXCEPT БЛОК для обработки ошибок скачивания, распаковки и подготовки ---
    except Exception as e:
        # Этот блок ловит ЛЮБУЮ ошибку, которая произошла
        # на этапах скачивания, распаковки, проверки или перемещения.
        log_message(f"Ошибка в процессе подготовки дистрибутива: {e}", level="ERROR")
        # Статус уже обновлен в более специфичных блоках except (если они были), но можно обновить еще раз
        update_status(f"Ошибка подготовки дистрибутива: {e}", level="ERROR")

        # --- ОШИБКА: Выполняем очистку временных папок и локальной папки дистрибутива ---
        # Удаляем временную папку распаковки, если она была создана
        if os.path.exists(temp_extract_path):
             try:
                 # ignore_errors=True полезно, если папка частично удалена или заблокирована
                 shutil.rmtree(temp_extract_path, ignore_errors=True)
                 log_message("Временная папка распаковки очищена после ошибки.")
             except Exception as temp_e:
                 log_message(f"Ошибка при очистке временной папки распаковки '{temp_extract_path}' после основной ошибки: {temp_e}", level="WARNING")

        # Удаляем временный архив, если он был создан (частично или полностью)
        # Проверяем флаг temp_archive_path_exists, чтобы не пытаться удалить, если скачивание даже не началось
        if temp_archive_path_exists and os.path.exists(temp_archive_path):
             try:
                 os.remove(temp_archive_path)
                 log_message("Временный архив удален после ошибки.")
             except Exception as temp_e:
                 log_message(f"Ошибка при очистке временного архива '{temp_archive_path}' после основной ошибки: {temp_e}", level="WARNING")

        # Очищаем локальную папку дистрибутива, если она была создана, но подготовка не завершилась успешно
        # Это важно, чтобы при следующей попытке не использовать неполный или некорректный дистрибутив.
        if os.path.exists(local_installer_path):
             try:
                 shutil.rmtree(local_installer_path, ignore_errors=True)
                 log_message("Локальная папка дистрибутива очищена после ошибки подготовки.")
             except Exception as cleanup_e:
                  log_message(f"Ошибка при очистке локальной папки дистрибутива '{local_installer_path}' после ошибки: {cleanup_e}", level="WARNING")


        update_progress_callback(0) # Прогресс 0% при ошибке
        return None


def get_appdata_path(vendor, app_type, sanitized_target, version_raw=None):
    """Определяет правильный путь к временной папке кэша BackOffice в AppData."""
    log_message(f"DEBUG: Определение пути AppData для производителя '{vendor}', типа '{app_type}', адреса '{sanitized_target}' и версии '{version_raw}'")
    app_data_root = os.getenv('APPDATA')
    if not app_data_root:
        log_message("Переменная окружения APPDATA не найдена.", level="ERROR")
        return None

    vendor_folder = "iiko" # Папка по умолчанию
    if vendor.lower() == "syrve":
        # Для Syrve версий 9+ используется папка 'Syrve', для более старых - 'iiko'
        try:
            if version_raw:
                version_parts = version_raw.split('.')
                if version_parts and len(version_parts) > 0:
                    major_version_str = re.match(r"^\d+", version_parts[0]) # Берем только цифры из первой части
                    if major_version_str:
                        major_version = int(major_version_str.group(0))
                        if major_version >= 9:
                            vendor_folder = "Syrve"
                            log_message(f"DEBUG: Версия Syrve >= 9 ('{version_raw}'). Используется папка Syrve в AppData.")
                        else:
                             log_message(f"DEBUG: Версия Syrve < 9 ('{version_raw}'). Используется папка iiko в AppData.")
                    else:
                         log_message(f"DEBUG: Не удалось извлечь основную версию из '{version_raw}'. Используется папка iiko.", level="WARNING")
                else:
                     log_message(f"DEBUG: Строка версии '{version_raw}' пуста или некорректна. Используется папка iiko.", level="WARNING")
            else:
                 log_message("DEBUG: Версия не предоставлена для логики папки Syrve. Используется папка iiko в AppData.", level="WARNING")

        except (ValueError, IndexError) as e:
             log_message(f"Ошибка парсинга основной версии из '{version_raw}' для логики папки Syrve: {e}. Используется папка iiko.", level="WARNING")


    # Определяем промежуточную папку: Rms или Chain
    intermediate_folder = "Rms" if "rms" in app_type.lower() else "Chain"

    # Объединяем части пути
    backoffice_temp_dir = os.path.join(app_data_root, vendor_folder, intermediate_folder, sanitized_target)
    log_message(f"Определен полный путь временной папки кэша: '{backoffice_temp_dir}' (Папка производителя в AppData: '{vendor_folder}')")
    return backoffice_temp_dir


def wait_for_file(filepath, timeout_sec, check_interval_ms, update_status, update_progress_step_callback):
    """Ожидает появления файла с таймаутом и обновлением прогресса."""
    log_message(f"Ожидание появления файла: '{filepath}'")
    update_status(f"Ожидание файла: '{os.path.basename(filepath)}'...")

    max_attempts = int(timeout_sec * 1000 / check_interval_ms)
    file_found = False

    # Прогресс для этого этапа: 0% -> 50% на ожидание файла
    step_progress_base = 0.0
    step_progress_range = 0.5

    for attempt in range(max_attempts):
        if os.path.exists(filepath):
            file_found = True
            log_message("Файл найден!")
            update_status("Файл конфигурации найден.")
            update_progress_step_callback(step_progress_base + step_progress_range) # Прогресс 50% для этапа "Ожидание файла"
            break
        time.sleep(check_interval_ms / 1000)
        # Обновляем прогресс немного во время ожидания
        # Прогресс за один цикл: (1 / max_attempts) * step_progress_range
        if max_attempts > 0: # Избегаем деления на ноль
            progress_increment_per_attempt = step_progress_range / max_attempts
            update_progress_step_callback(step_progress_base + (attempt + 1) * progress_increment_per_attempt)


    if not file_found:
        log_message(f"Таймаут ожидания файла конфигурации '{filepath}' ({timeout_sec} сек).", level="ERROR")
        update_status("Таймаут ожидания файла.", level="ERROR")
        return False

    # Ожидание содержимого файла (не пустой/не заблокирован)
    log_message(f"Ожидание содержимого в файле: '{filepath}'")
    update_status("Ожидание содержимого файла...")
    content_wait_timeout_sec = 10 # Таймаут ожидания содержимого
    content_check_interval_ms = 50
    max_content_attempts = int(content_wait_timeout_sec * 1000 / content_check_interval_ms)
    content_found = False

    # Прогресс для этого этапа: 50% -> 100% на ожидание содержимого
    step_progress_base = 0.5
    step_progress_range = 0.5

    for attempt in range(max_content_attempts):
        try:
            # Пытаемся прочитать небольшую часть, чтобы убедиться в доступности и наличии хоть чего-то
            # Используем 'r' и UTF-8, т.к. это текстовый XML файл
            with open(filepath, 'r', encoding='utf-8') as f:
                content_preview = f.read(100) # Читаем первые 100 символов
                # Базовая проверка на наличие XML-подобного содержимого (наличие '<')
                if content_preview and '<' in content_preview:
                    content_found = True
                    log_message("Содержимое файла обнаружено.")
                    update_status("Содержимое файла обнаружено.")
                    update_progress_step_callback(step_progress_base + step_progress_range) # Прогресс 100% для этапа "Ожидание файла"
                    break
        except Exception as e:
            # Ошибка при чтении (файл заблокирован или пуст?)
            log_message(f"DEBUG: Ошибка при чтении превью файла '{filepath}': {e}", level="VERBOSE") # Используем VERBOSE для отладочных логов

        time.sleep(content_check_interval_ms / 1000)
        # Обновляем прогресс немного во время ожидания содержимого
        # Прогресс за один цикл: (1 / max_content_attempts) * step_progress_range
        if max_content_attempts > 0: # Избегаем деления на ноль
            progress_increment_per_attempt = step_progress_range / max_content_attempts
            update_progress_step_callback(step_progress_base + (attempt + 1) * progress_increment_per_attempt)


    if not content_found:
        log_message(f"Таймаут ожидания содержимого в файле конфигурации '{filepath}' ({content_wait_timeout_sec} сек). Файл пуст или некорректен?", level="ERROR")
        update_status("Таймаут ожидания содержимого файла.", level="ERROR")
        return False

    return True


def edit_config_file(filepath, target_url_or_ip, target_port, config_protocol, update_status):
    """Редактирует файл backclient.config.xml."""
    log_message(f"Редактирование файла конфигурации: '{filepath}'")
    update_status("Редактирование файла конфигурации...")

    try:
        # Ожидаем короткое время после остановки процесса, чтобы файл точно был доступен
        time.sleep(0.5)

        # Используем ElementTree для парсинга и редактирования XML
        # parse() автоматически определяет кодировку, но явно укажем для надежности
        tree = ET.parse(filepath)
        root = tree.getroot()

        # Находим узел ServersList в любом месте дерева
        servers_list_node = root.find('.//ServersList')

        if servers_list_node is not None:
            log_message("DEBUG: Узел ServersList найден.")

            # Находим и обновляем ServerAddr
            server_addr_node = servers_list_node.find('ServerAddr')
            if server_addr_node is not None:
                server_addr_node.text = target_url_or_ip
                log_message(f"  Обновлен ServerAddr на '{target_url_or_ip}'")
            else:
                log_message("  Узел ServerAddr не найден под ServersList. Не удалось обновить.", level="WARNING")

            # Находим и обновляем Protocol
            protocol_node = servers_list_node.find('Protocol')
            if protocol_node is not None:
                protocol_node.text = config_protocol
                log_message(f"  Обновлен Protocol на '{config_protocol}'")
            else:
                log_message("  Узел Protocol не найден под ServersList. Не удалось обновить.", level="WARNING")

            # Находим и обновляем Port
            port_node = servers_list_node.find('Port')
            if port_node is not None:
                port_node.text = str(target_port) # Порт должен быть строкой в XML
                log_message(f"  Обновлен Port на '{target_port}'")
            else:
                log_message("  Узел Port не найден под ServersList. Не удалось обновить.", level="WARNING")

            # Сохраняем измененный XML обратно в файл
            # xml_declaration=True добавляет <?xml version='1.0' encoding='utf-8'?>
            # pretty_print=True (доступно в lxml, не в ET) помогло бы форматировать, но не критично
            tree.write(filepath, encoding='utf-8', xml_declaration=True)
            log_message("Файл конфигурации успешно обновлен.")
            update_status("Файл конфигурации успешно обновлен.")
            return True

        else:
            log_message(f"Узел ServersList не найден в файле конфигурации '{filepath}'. Не удалось отредактировать.", level="ERROR")
            update_status("Ошибка: Узел ServersList не найден в конфиге.", level="ERROR")
            return False

    except FileNotFoundError:
        log_message(f"Ошибка: Файл конфигурации не найден для редактирования: '{filepath}'", level="ERROR")
        update_status("Ошибка: Файл конфига не найден.", level="ERROR")
        return False
    except ET.ParseError as e:
        log_message(f"Ошибка парсинга XML файла '{filepath}': {e}", level="ERROR")
        update_status(f"Ошибка парсинга XML: {e}", level="ERROR")
        return False
    except Exception as e:
        log_message(f"Произошла ошибка при работе с файлом конфигурации '{filepath}': {e}", level="ERROR")
        update_status(f"Ошибка редактирования конфига: {e}", level="ERROR")
        return False

def stop_process_by_pid(pid):
    """Останавливает процесс по его PID (для Windows)."""
    if pid is None:
        log_message("PID процесса BackOffice не известен, пропуск остановки.", level="WARNING")
        return False # Не удалось остановить

    log_message(f"Попытка остановить процесс BackOffice (PID: {pid})...")
    try:
        # На Windows taskkill - надежный способ принудительно завершить процесс
        # /F - принудительно
        # /T - завершить дочерние процессы (не обязательно, но безопасно)
        # creationflags=subprocess.CREATE_NO_WINDOW - предотвращает появление консольного окна taskkill
        result = subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], check=True, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
        log_message(f"Процесс BackOffice (PID: {pid}) успешно остановлен. stdout: {result.stdout.strip()}")
        return True # Успешно остановлен
    except subprocess.CalledProcessError as e:
        # taskkill вернет ненулевой код, если процесс не найден (уже завершен) или нет прав
        log_message(f"Ошибка при остановке процесса BackOffice (PID: {pid}). Код выхода: {e.returncode}, stdout: {e.stdout.strip()}, stderr: {e.stderr.strip()}", level="WARNING")
        # Если код 128, это означает "процесс не найден". Считаем это успехом (процесс не запущен)
        if e.returncode == 128:
             log_message(f"Процесс (PID: {pid}) не найден, вероятно, уже завершен.", level="INFO")
             return True # Считаем, что процесс остановлен (или уже был)
        return False # Не удалось остановить по другой причине
    except FileNotFoundError:
        log_message("Команда 'taskkill' не найдена. Убедитесь, что вы на Windows.", level="ERROR")
        return False # Не удалось остановить
    except Exception as e:
        log_message(f"Неизвестная ошибка при остановке процесса BackOffice (PID: {pid}): {e}", level="ERROR")
        return False # Не удалось остановить


# --- Класс GUI ---

class BackOfficeLauncherGUI:
    def __init__(self, root):
        self.root = root
        root.title("BackOffice Launcher")
        # Размер окна оставим прежним, т.к. текстовое поле будет растягиваться
        root.geometry("600x290") # Ширина x Высота
        root.resizable(False, True) # Теперь окно может изменять размер

        # Загружаем конфигурацию сразу при старте
        self.config = load_config()
        # Устанавливаем глобальную переменную для логирования отладки из загруженной конфигурации
        global global_debug_logging # Объявляем использование глобальной переменной
        global_debug_logging = get_config_value(self.config, 'Settings', 'DebugLogging', default=False, type_cast=bool)
        log_message(f"Отладочное логирование включено: {global_debug_logging}", level="INFO")
        # Проверяем доступность ctypes для WinAPI на старте
        if os.name != 'nt':
            log_message(f"Запуск не на Windows. Функция get_file_company_name через WinAPI будет недоступна. Текущая ОС: {os.name}", level="WARNING")


        self.frame = ttk.Frame(root, padding="15") # Увеличиваем отступы
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Настраиваем растягивание столбцов и строк корневого окна
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Настраиваем 4 колонки внутри фрейма:
        # Колонка 0: Метка (не растягивается)
        # Колонка 1: Поле ввода (растягивается)
        # Колонка 2: Кнопка (не растягивается)
        # Колонка 3: Кнопка (не растягивается)
        self.frame.columnconfigure(1, weight=1) # Колонка для поля ввода растягивается
        self.frame.columnconfigure(0, weight=0) # Метка
        self.frame.columnconfigure(2, weight=0) # Кнопка Check - левая из пары
        self.frame.columnconfigure(3, weight=0) # Кнопка Launch/Paste - правая из пары

        # Метка "Введите URL или IP:порт:" - Ряд 0, Колонка 0
        ttk.Label(self.frame, text="Введите URL или IP:порт:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)

        # Поле ввода URL/IP - Ряд 0, Колонки 1-2 (чтобы быть шире)
        self.target_entry = ttk.Entry(self.frame)
        self.target_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=5) # columnspan=2
        self.target_entry.focus()
        self.target_entry.bind("<Return>", self.start_launch) # Привязываем клавишу Enter
        # Привязываем виртуальное событие <<Paste>> к нашему обработчику
        self.target_entry.bind("<<Paste>>", self._on_paste)


        # Кнопка "Paste" - Ряд 0, Колонка 3
        self.paste_button = ttk.Button(self.frame, text="Paste", command=self.paste_from_clipboard)
        self.paste_button.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)


        # Метка статуса - ПЕРЕНОСИМ В РЯД 1, Колонки 0-1
        # Она должна занимать место слева от кнопок Проверить/Запустить
        self.status_label = ttk.Label(self.frame, text="Ожидание ввода...", wraplength=400) # wraplength можно скорректировать
        self.status_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5, padx=5) # columnspan=2


        # Кнопки действий - ПЕРЕНОСИМ В РЯД 1
        # Кнопка "Проверить" - Ряд 1, Колонка 2
        self.check_button = ttk.Button(self.frame, text="Проверить", command=self.start_check)
        self.check_button.grid(row=1, column=2, sticky=tk.E, pady=5, padx=5) # sticky=tk.E для выравнивания по правому краю

        # Кнопка "Запустить" - Ряд 1, Колонка 3
        self.launch_button = ttk.Button(self.frame, text="Запустить", command=self.start_launch)
        self.launch_button.grid(row=1, column=3, sticky=tk.E, pady=5, padx=5) # sticky=tk.E


        # Прогресс бар - ПЕРЕНОСИМ В РЯД 2, Колонки 0-3
        self.progress_bar = ttk.Progressbar(self.frame, orient='horizontal', mode='determinate')
        self.progress_bar.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5, padx=5) # columnspan=4


        # Текстовое поле для вывода JSON - ПЕРЕНОСИМ В РЯД 3, Колонки 0-3
        # Используем tk.Text вместо scrolledtext.ScrolledText
        # Удаляем параметры width и height, т.к. размер будет определяться grid и растягиванием
        self.json_output_text = tk.Text(self.frame, wrap=tk.WORD, state=tk.DISABLED) # wrap=tk.WORD переносит по словам
        self.json_output_text.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=5)
        # Настраиваем растягивание ряда с текстовым полем
        self.frame.rowconfigure(3, weight=1) # Ряд с текстовым полем растягивается


        # Переменные для хранения информации о процессе и прогрессе
        self.backoffice_process = None
        self.current_step_progress = 0.0 # Прогресс внутри текущего основного шага (от 0.0 до 1.0)
        self.progress_base = 0.0 # Начальное значение прогресса для текущего шага (0-100)
        self.next_step_base = 0.0 # Начальное значение прогресса для следующего шага (0-100)

        # Словарь для сохранения данных между шагами, которые требуют пользовательского ввода
        self._launch_data = {}


    def update_status(self, message, level="INFO"):
        """Обновляет метку статуса GUI и выводит сообщение в лог."""
        # Обновление GUI должно происходить в основном потоке
        color = "black"
        if level == "ERROR":
            color = "red"
        elif level == "WARNING":
            color = "orange"
        self.root.after(0, lambda: self.status_label.config(text=message, foreground=color))
        log_message(message, level)

    def update_progress(self, value):
        """Обновляет значение прогресс бара GUI (0-100)."""
        # Обновление GUI должно происходить в основном потоке
        # Убедимся, что значение находится в диапазоне [0, 100]
        safe_value = max(0, min(100, value))
        self.root.after(0, lambda: self.progress_bar.config(value=safe_value))

    def reset_progress(self):
         """Сбрасывает прогресс бар на 0."""
         self.update_progress(0)
         self.current_step_progress = 0.0
         self.progress_base = 0.0
         self.next_step_base = 0.0

    def set_progress_step_bounds(self, base_value, next_base_value):
        """Устанавливает границы прогресса (0-100) для текущего шага."""
        self.progress_base = base_value
        self.next_step_base = next_base_value
        self.current_step_progress = 0.0 # Сбрасываем прогресс внутри шага
        log_message(f"DEBUG: Установлены границы прогресса шага: {self.progress_base}% - {self.next_step_base}%")


    def update_progress_step(self, increment_factor=0.0):
        """Обновляет прогресс внутри текущего шага.
           increment_factor - это доля от 0 до 1, насколько завершен текущий шаг.
           Если increment_factor = 1.0, это означает, что текущий шаг полностью завершен.
           Если increment_factor = 0.5, это означает, что текущий шаг завершен на 50%.
           Эта функция не добавляет к текущему self.current_step_progress, а устанавливает его долю.
        """
        # Устанавливаем долю завершенности текущего шага
        self.current_step_progress = max(0.0, min(1.0, increment_factor))

        # Вычисляем общий прогресс (0-100)
        step_range = self.next_step_base - self.progress_base
        total_progress = self.progress_base + (self.current_step_progress * step_range)
        self.update_progress(total_progress)
        log_message(f"DEBUG: Прогресс обновлен. Шаг: {self.progress_base:.1f}-{self.next_step_base:.1f}%, Внутри шага: {self.current_step_progress*100:.1f}%, Общий: {total_progress:.1f}%")


    def start_launch(self, event=None):
        """Запускает процесс запуска BackOffice в отдельном потоке."""
        target_string = self.target_entry.get().strip()
        if not target_string:
            self.update_status("Введите URL или IP:порт.", level="WARNING")
            return

        # Отключаем ввод и кнопку на время выполнения
        self.launch_button.config(state=tk.DISABLED)
        self.target_entry.config(state=tk.DISABLED)
        self.reset_progress()
        self.update_status("Запуск процесса...")

        # Очищаем данные предыдущего запуска
        self._launch_data = {}
        self._launch_data['target_string'] = target_string

        # Используем поток, чтобы не блокировать GUI
        self.launch_thread = threading.Thread(target=self._launch_step1_parse, args=(target_string,))
        self.launch_thread.daemon = True # Поток завершится при закрытии приложения
        self.launch_thread.start()

    def paste_from_clipboard(self):
        """Generates a paste event on the target entry."""
        # Генерируем виртуальное событие <<Paste>> на поле ввода.
        # Обработка вставки (получение из буфера, ограничение, вставка)
        # будет выполняться в методе _on_paste, привязанном к этому событию.
        log_message("Кнопка 'Paste' нажата. Генерация события <<Paste>>.", level="DEBUG")
        self.target_entry.event_generate("<<Paste>>")

    def start_check(self):
        """Starts the server check process in a separate thread."""
        target_string = self.target_entry.get().strip()
        if not target_string:
            self.update_status("Введите URL или IP:порт для проверки.", level="WARNING")
            return

        # Отключаем ввод и кнопки на время выполнения
        self.launch_button.config(state=tk.DISABLED)
        self.check_button.config(state=tk.DISABLED)
        self.target_entry.config(state=tk.DISABLED)

        # Очищаем текстовое поле
        self.update_text_area("")

        self.update_status("Выполнение проверки сервера...")
        self.reset_progress() # Сбрасываем прогресс для проверки (хотя она короткая)
        self.update_progress(0) # Убедимся, что прогресс на 0

        # Используем поток для выполнения проверки
        self.check_thread = threading.Thread(target=self.check_server_thread, args=(target_string,))
        self.check_thread.daemon = True # Поток завершится при закрытии приложения
        self.check_thread.start()

    def check_server_thread(self, target_string):
        """Performs the server check logic in a separate thread."""
        try:
            # Шаг проверки: Парсинг
            self.update_status("Парсинг адреса для проверки...")
            parsed_target = parse_target_string(target_string)
            if parsed_target is None or not parsed_target.get('UrlOrIp'):
                raise ValueError("Не удалось распарсить ввод или извлечь хост/IP.")

            target_url_or_ip = parsed_target['UrlOrIp']
            target_port = parsed_target['Port']

            # Шаг проверки: Формирование URL и запрос
            self.update_status(f"Запрос информации о сервере: {target_url_or_ip}:{target_port}...")

            # Логика формирования probe_url такая же, как в _launch_step2_httprequest
            probe_scheme = "http"
            if target_url_or_ip.lower().endswith(".iiko.it") or target_url_or_ip.lower().endswith(".syrve.online"):
                probe_scheme = "https"

            standard_http_port = 80
            standard_https_port = 443
            include_port_in_probe_url = True

            if probe_scheme == "http" and target_port == standard_http_port:
                 include_port_in_probe_url = False
            elif probe_scheme == "https" and target_port == standard_https_port:
                 include_port_in_probe_url = False

            probe_url = f"{probe_scheme}://{target_url_or_ip}"
            if include_port_in_probe_url:
                probe_url += f":{target_port}"
            probe_url += "/resto/getServerMonitoringInfo.jsp"

            log_message(f"URL для запроса информации о сервере (проверка): {probe_url}")

            http_timeout = get_config_value(self.config, 'Settings', 'HttpRequestTimeoutSec', default=15, type_cast=int)
            server_info = None
            try:
                response = requests.get(probe_url, timeout=http_timeout)
                response.raise_for_status()
                server_info = response.json()
            except requests.exceptions.Timeout:
                raise ConnectionError(f"Таймаут ({http_timeout} сек) при выполнении GET-запроса к '{probe_url}'")
            except requests.exceptions.ConnectionError as e:
                 raise ConnectionError(f"Ошибка подключения при выполнении GET-запроса к '{probe_url}': {e}")
            except requests.exceptions.RequestException as e:
                raise ConnectionError(f"Ошибка HTTP запроса к '{probe_url}': {e}")
            except Exception as e:
                 raise ConnectionError(f"Неожиданная ошибка при запросе к '{probe_url}': {e}")

            # Шаг проверки: Вывод результата
            self.update_status("Получен ответ от сервера. Вывод JSON...")
            log_message(f"Получен ответ от сервера: {server_info}", level="DEBUG")

            # Форматируем JSON для удобного чтения
            formatted_json = json.dumps(server_info, indent=4, ensure_ascii=False)
            self.update_text_area(formatted_json) # Выводим в текстовое поле

            self.update_status("Проверка завершена. Ответ сервера в поле ниже.", level="INFO")
            self.update_progress(100) # Прогресс 100% для проверки

        except Exception as e:
            # Обработка ошибок проверки
            log_message(f"Ошибка во время проверки сервера: {e}", level="ERROR")
            self.update_status(f"Ошибка проверки: {e}", level="ERROR")
            self.update_progress(0) # Сбрасываем прогресс при ошибке
            self.update_text_area(f"Ошибка при проверке сервера:\n{e}") # Выводим ошибку в текстовое поле


        finally:
            # Включаем ввод и кнопки обратно в основном потоке
            self.root.after(0, lambda: self.launch_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.check_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.target_entry.config(state=tk.NORMAL))
            self.reset_progress() # Убедимся, что прогресс сброшен

    def update_text_area(self, text):
        """Updates the text area safely from a thread."""
        # Обновление GUI должно происходить в основном потоке
        self.root.after(0, lambda: self._do_update_text_area(text))

    def _do_update_text_area(self, text):
        """Performs the actual text area update in the main thread."""
        self.json_output_text.config(state=tk.NORMAL) # Включаем редактирование
        self.json_output_text.delete('1.0', tk.END) # Очищаем
        self.json_output_text.insert(tk.END, text) # Вставляем новый текст
        self.json_output_text.config(state=tk.DISABLED) # Отключаем редактирование снова
        # Прокручиваем к началу, если текст не пустой
        if text:
            self.json_output_text.see('1.0')

    def _launch_step1_parse(self, target_string):
        """Шаг 1: Парсинг ввода (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(0, 5) # 0-5% на парсинг
            self.update_status("Парсинг введенного адреса...")
            parsed_target = parse_target_string(target_string)
            if parsed_target is None or not parsed_target.get('UrlOrIp'): # Проверяем, что хост не пустой
                raise ValueError("Не удалось распарсить ввод или извлечь хост/IP.")
            self.update_progress_step(1.0) # Шаг парсинга завершен

            self._launch_data['parsed_target'] = parsed_target

            # Переходим к следующему шагу
            self._launch_step2_httprequest(parsed_target)

        except Exception as e:
            # Перехватываем любые исключения в этом потоке и обрабатываем их
            self.handle_error(f"Произошла ошибка во время парсинга: {e}")

    def _on_paste(self, event=None):
        """Handles the <<Paste>> event, inserting limited clipboard content."""
        log_message("DEBUG: Событие <<Paste>> перехвачено.", level="DEBUG")
        try:
            # Получаем содержимое буфера обмена
            clipboard_content = self.root.clipboard_get()

            # Проверяем, является ли содержимое строкой и не пустое
            if isinstance(clipboard_content, str) and clipboard_content.strip():
                # Ограничиваем содержимое первыми 100 символами после удаления начальных/конечных пробелов
                truncated_content = clipboard_content.strip()[:100]

                # Очищаем поле ввода
                self.target_entry.delete('0', tk.END)
                log_message("DEBUG: Поле ввода очищено.", level="DEBUG")

                # Вставляем ограниченное содержимое в позицию курсора
                self.target_entry.insert(0, truncated_content)
                log_message(f"DEBUG: Вставлено содержимое (ограничено до 100 символов): '{truncated_content}'", level="DEBUG")

                # Очищаем статус после успешной вставки (если там было предупреждение)
                self.update_status("Ожидание ввода...")

            else:
                # Если содержимое не строка, пустое или состоит только из пробелов
                log_message("Буфер обмена пуст или содержит нетекстовые данные.", level="WARNING")
                self.update_status("Буфер обмена пуст или содержит нетекстовые данные.", level="WARNING")

        except tk.TclError:
            # Обрабатываем случаи, когда буфер обмена пуст или содержит нетекстовые данные (если strip/[:100] не сработали)
            # или если clipboard_get сам по себе вызвал ошибку (например, нет доступа к буферу)
            log_message("Не удалось получить содержимое из буфера обмена (TclError при clipboard_get).", level="WARNING")
            self.update_status("Буфер обмена пуст или содержит нетекстовые данные.", level="WARNING")
        except Exception as e:
            log_message(f"Неизвестная ошибка при обработке события <<Paste>>: {e}", level="ERROR")
            self.update_status(f"Ошибка при вставке: {e}", level="ERROR")

        # Важно: возвращаем 'break' чтобы остановить дальнейшую обработку события <<Paste>>
        # Иначе стандартная привязка tk вставит полное содержимое буфера после нашего кода
        return 'break'
    
    def _launch_step2_httprequest(self, parsed_target):
        """Шаг 2: Выполнение HTTP-запроса (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(5, 25) # 5-25% на HTTP запрос
            target_url_or_ip = parsed_target['UrlOrIp']
            target_port = parsed_target['Port']
            self.update_status(f"Выполнение GET-запроса к {target_url_or_ip}:{target_port}...")

            # Определяем URL запроса и схему
            # Используем HTTPS только для доменов iiko.it или syrve.online
            probe_scheme = "http"
            if target_url_or_ip.lower().endswith(".iiko.it") or target_url_or_ip.lower().endswith(".syrve.online"):
                probe_scheme = "https"
                log_message(f"DEBUG: Определена схема HTTPS для домена '{target_url_or_ip}'")
            else:
                 log_message(f"DEBUG: Определена схема HTTP для адреса '{target_url_or_ip}'")

            # Определяем, нужно ли включать порт в URL запроса
            standard_http_port = 80
            standard_https_port = 443
            include_port_in_probe_url = True

            if probe_scheme == "http" and target_port == standard_http_port:
                 include_port_in_probe_url = False
            elif probe_scheme == "https" and target_port == standard_https_port:
                 include_port_in_probe_url = False

            probe_url = f"{probe_scheme}://{target_url_or_ip}"
            if include_port_in_probe_url:
                probe_url += f":{target_port}"
            probe_url += "/resto/getServerMonitoringInfo.jsp"

            log_message(f"URL для запроса информации о сервере: {probe_url}")
            self._launch_data['probe_url'] = probe_url
            self._launch_data['config_protocol'] = probe_scheme # Сохраняем схему для config файла

            http_timeout = get_config_value(self.config, 'Settings', 'HttpRequestTimeoutSec', default=15, type_cast=int)
            response = None
            try:
                # Выполняем GET запрос
                response = requests.get(probe_url, timeout=http_timeout)
                response.raise_for_status() # Генерирует исключение для плохих кодов статуса (4xx или 5xx)
                server_info = response.json() # Парсим JSON ответ
                self.update_progress_step(1.0) # Шаг HTTP запроса завершен

                log_message(f"Получен ответ от сервера на шаге запуска: {server_info}", level="DEBUG")
                formatted_json = json.dumps(server_info, indent=4, ensure_ascii=False)
                self.update_text_area(formatted_json)

            except requests.exceptions.Timeout:
                raise ConnectionError(f"Таймаут ({http_timeout} сек) при выполнении GET-запроса к '{probe_url}'")
            except requests.exceptions.ConnectionError as e:
                 raise ConnectionError(f"Ошибка подключения при выполнении GET-запроса к '{probe_url}': {e}")
            except requests.exceptions.RequestException as e:
                raise ConnectionError(f"Ошибка HTTP запроса к '{probe_url}': {e}")
            except Exception as e:
                 raise ConnectionError(f"Неожиданная ошибка при запросе к '{probe_url}': {e}")

            self._launch_data['server_info'] = server_info

            # Переходим к следующему шагу
            self._launch_step3_process_response(server_info)

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время HTTP запроса: {e}")


    def _launch_step3_process_response(self, server_info):
        """Шаг 3: Обработка ответа и определение типа приложения (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(25, 40) # 25-40% на обработку ответа и определение типа
            self.update_status("Обработка ответа сервера...")

            edition = server_info.get("edition")
            version_raw = server_info.get("version")
            server_state = server_info.get("serverState")

            if None in [edition, version_raw, server_state]:
                 log_message(f"Полный ответ сервера: {server_info}", level="DEBUG")
                 raise ValueError("Ответ сервера не содержит ожидаемых ключей (edition, version, serverState).")

            self.update_status(f"Получены данные: Edition={edition}, Version={version_raw}, State={server_state}")
            log_message(f"Получены данные сервера: Edition='{edition}', Version='{version_raw}', ServerState='{server_state}'")

            self._launch_data['edition'] = edition
            self._launch_data['version_raw'] = version_raw
            self._launch_data['server_state'] = server_state

            # Определение типа приложения и производителя
            self.update_status("Определение типа приложения и производителя...")
            app_info = determine_app_type(self._launch_data['target_string'], edition)

            # Если app_info == None, значит, требуется выбор пользователя.
            # Мы не можем показывать диалог из этого потока.
            # Используем root.after, чтобы запланировать вызов функции в основном потоке.
            if app_info is None:
                 self.update_progress_step(0.5) # Прогресс 50% на этом шаге (до диалога)
                 log_message("Требуется выбор типа приложения пользователем.")
                 # Сохраняем данные для передачи в функцию диалога
                 self._launch_data['step_after_app_type'] = '_launch_step4_check_server_state' # Куда вернуться после выбора
                 self.root.after(0, self.ask_app_type) # Вызываем диалог в основном потоке
                 return # Завершаем выполнение этого потока, ожидая выбора пользователя

            # Если тип приложения определен автоматически, сохраняем и переходим к следующему шагу
            self._launch_data['app_type'] = app_info['AppType']
            self._launch_data['vendor'] = app_info['Vendor']
            self.update_status(f"Определен тип приложения: '{self._launch_data['app_type']}' (Производитель: '{self._launch_data['vendor']}')")
            log_message(f"Определен тип приложения: '{self._launch_data['app_type']}' (Производитель: '{self._launch_data['vendor']}')")
            self.update_progress_step(1.0) # Шаг определения типа завершен

            # Переходим к следующему шагу
            self._launch_step4_check_server_state(self._launch_data['server_state'])

        except Exception as e:
            self.handle_error(f"Произошла ошибка при обработке ответа сервера: {e}")


    def ask_app_type(self):
        """Показывает диалог для выбора типа приложения пользователем (вызывается в основном потоке)."""
        # Получаем данные, сохраненные в _launch_data
        target_string = self._launch_data['target_string']
        edition = self._launch_data['edition']
        # version_raw = self._launch_data['version_raw'] # Не нужны для диалога
        # server_state = self._launch_data['server_state'] # Не нужны для диалога

        vendor = "iiko"
        if "syrve" in target_string.lower():
            vendor = "Syrve"

        choices = [f"{vendor}RMS", f"{vendor}Chain"]
        # messagebox.askyesno возвращает True для Yes, False для No
        # Мы используем yesno для выбора между двумя опциями
        result = messagebox.askyesno("Выбор типа приложения",
                                        f"Не удалось автоматически определить тип RMS/Chain для производителя '{vendor}' по edition ('{edition}').\n"
                                        f"Выберите тип приложения:\n"
                                        f"  'Да' - {choices[0]}\n"
                                        f"  'Нет' - {choices[1]}",
                                        icon='question')

        selected_app_type = None
        if result is True: # Пользователь выбрал Yes (первая опция - RMS)
            selected_app_type = choices[0]
            log_message(f"Пользователь выбрал тип приложения: '{selected_app_type}' (Yes)")
        elif result is False: # Пользователь выбрал No (вторая опция - Chain)
            selected_app_type = choices[1]
            log_message(f"Пользователь выбрал тип приложения: '{selected_app_type}' (No)")
        else:
             log_message("Выбор типа приложения отменен пользователем (закрыл диалог).", level="WARNING")
             self.handle_error("Выбор типа приложения отменен пользователем.")
             return # Выходим, если выбор отменен

        # Определяем производителя снова на основе выбранного типа
        selected_vendor = "iiko" if "iiko" in selected_app_type.lower() else "Syrve"

        # Сохраняем выбранные данные
        self._launch_data['app_type'] = selected_app_type
        self._launch_data['vendor'] = selected_vendor

        # Возобновляем процесс запуска с выбранным типом приложения в НОВОМ потоке
        self.update_status(f"Продолжение запуска (выбран тип: {selected_app_type})...")
        # Начинаем новый поток с шага, который должен был идти после определения типа
        next_step_method_name = self._launch_data.get('step_after_app_type', '_launch_step4_check_server_state')
        next_step_method = getattr(self, next_step_method_name)

        # Устанавливаем прогресс так, как если бы шаг 3 завершился на 100%
        self.set_progress_step_bounds(25, 40) # Шаг 3
        self.update_progress_step(1.0) # Шаг 3 завершен

        # Теперь запускаем следующий шаг (Шаг 4) в новом потоке
        self.launch_thread = threading.Thread(target=next_step_method, args=(self._launch_data['server_state'],))
        self.launch_thread.daemon = True
        self.launch_thread.start()


    def _launch_step4_check_server_state(self, server_state):
        """Шаг 4: Проверка состояния сервера (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(40, 45) # 40-45% на проверку состояния сервера
            parsed_target = self._launch_data['parsed_target'] # Получаем parsed_target
            target_url_or_ip = parsed_target['UrlOrIp']

            if server_state != "STARTED_SUCCESSFULLY":
                log_message(f"Состояние сервера '{target_url_or_ip}' не 'STARTED_SUCCESSFULLY', текущее состояние: '{server_state}'.", level="WARNING")
                self.update_status(f"Сервер в состоянии: '{server_state}'. Продолжить?", level="WARNING")
                 # Требуется подтверждение пользователя. Планируем вызов диалога в основном потоке.
                self.update_progress_step(0.5) # Прогресс 50% на этом шаге (до диалога)
                # Сохраняем данные для передачи в функцию диалога
                self._launch_data['step_after_server_state_confirm'] = '_launch_step5_format_version' # Куда вернуться после подтверждения
                self.root.after(0, self.ask_continue_on_server_state) # Вызываем диалог в основном потоке
                return # Завершаем выполнение этого потока, ожидая подтверждения

            self.update_status("Состояние сервера OK.")
            self.update_progress_step(1.0) # Шаг проверки состояния завершен

            # Если все проверки пройдены, переходим к следующему шагу
            self._launch_step5_format_version()

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время проверки состояния сервера: {e}")


    def ask_continue_on_server_state(self):
        """Показывает диалог для подтверждения продолжения при неидеальном состоянии сервера (вызывается в основном потоке)."""
        # Получаем данные, сохраненные в _launch_data
        parsed_target = self._launch_data['parsed_target']
        target_url_or_ip = parsed_target['UrlOrIp']
        server_state = self._launch_data['server_state']

        result = messagebox.askyesno("Состояние сервера",
                                        f"Состояние сервера '{target_url_or_ip}' не 'STARTED_SUCCESSFULLY', текущее состояние: '{server_state}'.\n"
                                        f"Продолжить запуск BackOffice?",
                                        icon='warning')

        if result is True: # Пользователь выбрал Yes
            log_message("Пользователь подтвердил продолжение запуска.")
            self.update_status("Продолжение запуска по запросу пользователя.")
            # Возобновляем процесс запуска в НОВОМ потоке
            # Начинаем новый поток с шага, который должен был идти после подтверждения
            next_step_method_name = self._launch_data.get('step_after_server_state_confirm', '_launch_step5_format_version')
            next_step_method = getattr(self, next_step_method_name)

            # Устанавливаем прогресс так, как если бы шаг 4 завершился на 100%
            self.set_progress_step_bounds(40, 45) # Шаг 4
            self.update_progress_step(1.0) # Шаг 4 завершен

            # Теперь запускаем следующий шаг (Шаг 5) в новом потоке
            self.launch_thread = threading.Thread(target=next_step_method)
            self.launch_thread.daemon = True
            self.launch_thread.start()

        else: # Пользователь выбрал No или закрыл диалог
            log_message("Запуск отменен по запросу пользователя.", level="INFO")
            self.handle_error("Запуск отменен по запросу пользователя.") # Используем handle_error для сброса состояния GUI


    def _launch_step5_format_version(self):
        """Шаг 5: Форматирование версии (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(45, 50) # 45-50% на форматирование версии
            version_raw = self._launch_data['version_raw']
            self.update_status("Форматирование версии...")
            version_formatted = format_version(version_raw)
            log_message(f"Форматированная версия: '{version_formatted}'")
            self.update_status(f"Форматированная версия: {version_formatted}")
            self.update_progress_step(1.0) # Шаг форматирования завершен

            self._launch_data['version_formatted'] = version_formatted

            # Переходим к следующему шагу
            self._launch_step6_get_installer_name(self._launch_data['app_type'], version_formatted)

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время форматирования версии: {e}")


    def _launch_step6_get_installer_name(self, app_type, version_formatted):
        """Шаг 6: Определение ожидаемого имени каталога дистрибутива (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(50, 55) # 50-55% на имя дистрибутива
            self.update_status("Определение имени дистрибутива...")
            # Передаем объект конфигурации
            expected_installer_name = get_expected_installer_name(self.config, app_type, version_formatted)
            if expected_installer_name is None:
                # Ошибка уже залогирована внутри get_expected_installer_name
                raise ValueError("Ошибка формирования имени дистрибутива.")
            log_message(f"Ожидаемое имя каталога дистрибутива: '{expected_installer_name}'")
            self.update_status(f"Ожидаемое имя дистрибутива: {expected_installer_name}")
            self.update_progress_step(1.0) # Шаг определения имени завершен

            self._launch_data['expected_installer_name'] = expected_installer_name

            # Переходим к следующему шагу
            # Передаем объект конфигурации
            self._launch_step7_find_or_download_installer(self.config, expected_installer_name, self._launch_data['vendor'])

        except Exception as e:
            self.handle_error(f"Произошла ошибка при определении имени дистрибутива: {e}")


    def _launch_step7_find_or_download_installer(self, config, expected_installer_name, vendor):
        """Шаг 7: Поиск или скачивание дистрибутива (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(55, 85) # 55-85% на поиск/скачивание/распаковку (30%)

            # Вызываем функцию поиска/скачивания, передавая колбэки для обновления GUI и объект конфига
            installer_path = find_or_download_installer(
                config, # 1. config
                self._launch_data['app_type'], # 2. app_type
                self._launch_data['version_formatted'], # 3. version_formatted
                vendor, # 4. vendor (уже есть в аргументах метода)
                self.update_status, # 5. update_status
                self.update_progress # 6. update_progress_callback
            )

            if installer_path is None:
                # Ошибка уже была залогирована и статус обновлен внутри find_or_download_installer
                raise FileNotFoundError("Не удалось найти или подготовить дистрибутив.") # Перебрасываем ошибку

            # Прогресс должен быть обновлен до 100% внутри find_or_download_installer при успехе
            log_message(f"Каталог дистрибутива готов: '{installer_path}'")
            self.update_status(f"Каталог дистрибутива готов: {os.path.basename(installer_path)}")

            self._launch_data['installer_path'] = installer_path

            # Переходим к следующему шагу
            # Передаем нужные данные из _launch_data
            self._launch_step8_appdata_cleanup(self._launch_data['parsed_target']['UrlOrIp'], self._launch_data['vendor'], self._launch_data['app_type'], self._launch_data['version_raw'])

        except Exception as e:
            # Этот except блок ловит ошибки, которые были переброшены (raise)
            # из find_or_download_installer при неудаче.
            # find_or_download_installer уже выполнила очистку и обновила статус/прогресс при ошибке.
            # Здесь просто обрабатываем ошибку для GUI потока.
            self.handle_error(f"Произошла ошибка при поиске или скачивании дистрибутива: {e}")
            # update_status и update_progress(0) уже вызваны в find_or_download_installer при ошибке.
            # handle_error также обновит статус и сбросит прогресс.


    def _launch_step8_appdata_cleanup(self, target_url_or_ip, vendor, app_type, version_raw):
        """Шаг 8: Определение пути AppData и очистка (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(85, 88) # 85-88% на AppData и очистку
            # Санитизируем адрес для пути в AppData
            sanitized_target = sanitize_for_path(target_url_or_ip)
            log_message(f"Санитизированный адрес для пути AppData: '{sanitized_target}'")

            # Определяем путь к временной папке в AppData
            backoffice_temp_dir = get_appdata_path(vendor, app_type, sanitized_target, version_raw)
            if backoffice_temp_dir is None:
                 raise EnvironmentError("Не удалось определить путь временной папки кэша.")

            self.update_status(f"Ожидаемый путь временной папки кэша: {backoffice_temp_dir}")
            self._launch_data['backoffice_temp_dir'] = backoffice_temp_dir
            self._launch_data['sanitized_target'] = sanitized_target

            # Очищаем старый кэш в AppData, если существует
            if os.path.exists(backoffice_temp_dir):
                self.update_status(f"Очистка существующей временной папки кэша...")
                log_message(f"Очистка существующей временной папки кэша: '{backoffice_temp_dir}'")
                try:
                    shutil.rmtree(backoffice_temp_dir, ignore_errors=True)
                    log_message("Временная папка успешно удалена.")
                    self.update_status("Временная папка кэша очищена.")
                except Exception as e:
                    log_message(f"Ошибка при удалении временной папки '{backoffice_temp_dir}': {e}", level="ERROR")
                    self.update_status("Ошибка при удалении временной папки.", level="WARNING") # Предупреждение, не критическая ошибка
            else:
                log_message("Временная папка кэша не найдена. Удаление не требуется.")
                self.update_status("Временная папка кэша не найдена.")
            self.update_progress_step(1.0) # Шаг AppData и очистки завершен

            # Переходим к следующему шагу
            self._launch_step9_first_run()

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время определения пути AppData или очистки: {e}")


    def _launch_step9_first_run(self):
        """Шаг 9: Первый запуск BackOffice.exe (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(88, 90) # 88-90% на первый запуск
            installer_path = self._launch_data['installer_path']
            sanitized_target = self._launch_data['sanitized_target']

            backoffice_exe_path = os.path.join(installer_path, "BackOffice.exe")
            if not os.path.exists(backoffice_exe_path):
                 raise FileNotFoundError(f"Файл BackOffice.exe не найден в каталоге дистрибутива: '{backoffice_exe_path}'.")

            # Формируем аргументы для BackOffice.exe
            backoffice_args = f"/AdditionalTmpFolder=\"{sanitized_target}\""
            log_message(f"Первый запуск BackOffice.exe: '{backoffice_exe_path}' с аргументами: '{backoffice_args}'")
            self.update_status("Первый запуск BackOffice.exe...")

            try:
                # Используем subprocess.Popen для асинхронного запуска и получения объекта процесса
                # shell=True может быть опасен с пользовательским вводом, но здесь $sanitized_target уже очищен.
                # Альтернатива без shell=True сложнее с аргументами с пробелами и кавычками.
                # cwd устанавливает рабочий каталог.
                # startupinfo для скрытия окна консоли в Windows
                startupinfo = None
                if os.name == 'nt': # Проверяем, что мы на Windows
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE # Скрыть окно

                self.backoffice_process = subprocess.Popen(
                    [backoffice_exe_path, backoffice_args],
                    cwd=installer_path,
                    shell=True, # Используем shell=True для корректной обработки аргументов
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0 # Не создавать консольное окно на Windows
                )
                log_message(f"BackOffice.exe успешно запущен (первый раз, PID: {self.backoffice_process.pid}).")
                self.update_status(f"BackOffice.exe запущен (PID: {self.backoffice_process.pid}).")
                self.update_progress_step(1.0) # Шаг первого запуска завершен

                # Сохраняем информацию о процессе и аргументах
                self._launch_data['backoffice_process'] = self.backoffice_process
                self._launch_data['backoffice_exe_path'] = backoffice_exe_path
                self._launch_data['backoffice_args'] = backoffice_args

            except FileNotFoundError:
                 raise FileNotFoundError(f"Не удалось найти исполняемый файл BackOffice.exe: '{backoffice_exe_path}'. Убедитесь в правильности пути.")
            except Exception as e:
                 raise RuntimeError(f"Ошибка при первом запуске BackOffice.exe: {e}")


            # Переходим к следующему шагу
            self._launch_step10_wait_edit_config()

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время первого запуска BackOffice: {e}")


    def _launch_step10_wait_edit_config(self):
        """Шаг 10: Ожидание и редактирование файла backclient.config (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(90, 98) # 90-98% на ожидание и редактирование конфига (8%)
            backoffice_temp_dir = self._launch_data['backoffice_temp_dir']
            parsed_target = self._launch_data['parsed_target']
            target_url_or_ip = parsed_target['UrlOrIp']
            target_port = parsed_target['Port']
            config_protocol = self._launch_data['config_protocol'] # Протокол, определенный на шаге 2

            config_file_path = os.path.join(backoffice_temp_dir, "config", "backclient.config.xml")
            config_wait_timeout = get_config_value(self.config, 'Settings', 'ConfigFileWaitTimeoutSec', default=60, type_cast=int)
            config_check_interval = get_config_value(self.config, 'Settings', 'ConfigFileCheckIntervalMs', default=100, type_cast=int)

            # Ожидаем появления файла конфигурации
            # update_progress_step_callback передается в wait_for_file, чтобы обновлять прогресс внутри этого шага
            # Прогресс ожидания файла и содержимого внутри wait_for_file будет от 0.0 до 1.0
            if not wait_for_file(config_file_path, config_wait_timeout, config_check_interval, self.update_status, lambda p: self.update_progress_step(p)):
                 # Если wait_for_file вернул False, это таймаут или ошибка ожидания содержимого.
                 # Ошибка уже была залогирована и статус обновлен внутри wait_for_file.
                 if self._launch_data.get('backoffice_process') and self._launch_data['backoffice_process'].poll() is None: # Проверяем, запущен ли процесс
                     self.update_status("Таймаут ожидания файла конфига. Попытка остановить BackOffice...", level="WARNING")
                     stop_process_by_pid(self._launch_data['backoffice_process'].pid) # Останавливаем процесс
                 raise TimeoutError(f"Таймаут ожидания файла конфигурации '{config_file_path}'.")

            # Файл найден. Останавливаем процесс BackOffice перед редактированием.
            current_process = self._launch_data.get('backoffice_process')
            if current_process and current_process.poll() is None: # Проверяем, запущен ли процесс
                self.update_status("Файл конфигурации найден. Остановка процесса BackOffice для редактирования...")
                log_message(f"Остановка процесса BackOffice (PID: {current_process.pid}) для редактирования файла.")
                # stop_process_by_pid возвращает True, если успешно или процесс не найден
                stop_process_by_pid(current_process.pid)
                # Даем немного времени процессу на завершение и освобождение файла
                time.sleep(1.0)
                self._launch_data['backoffice_process'] = None # Сбрасываем объект процесса в данных запуска
                self.backoffice_process = None # Сбрасываем объект процесса в GUI классе


            # Редактируем файл конфигурации
            self.update_status("Редактирование файла конфигурации...")
            # edit_config_file не обновляет прогресс внутри себя, просто завершает шаг
            if not edit_config_file(config_file_path, target_url_or_ip, target_port, config_protocol, self.update_status):
                 # Ошибка уже была залогирована и статус обновлен внутри edit_config_file
                 raise RuntimeError(f"Не удалось отредактировать файл конфигурации '{config_file_path}'.")

            self.update_progress_step(1.0) # Шаг редактирования завершен (достигает 98%)

            # Переходим к следующему шагу
            self._launch_step11_restart()

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время ожидания или редактирования конфига: {e}")


    def _launch_step11_restart(self):
        """Шаг 11: Перезапуск BackOffice.exe (выполняется в потоке)."""
        try:
            self.set_progress_step_bounds(98, 100) # 98-100% на перезапуск (2%)
            installer_path = self._launch_data['installer_path']
            backoffice_exe_path = self._launch_data['backoffice_exe_path']
            backoffice_args = self._launch_data['backoffice_args']

            self.update_status("Перезапуск BackOffice.exe...")
            log_message(f"Перезапуск BackOffice.exe: '{backoffice_exe_path}' с аргументами: '{backoffice_args}'")

            try:
                 # Запускаем BackOffice.exe снова с теми же аргументами
                 startupinfo = None
                 if os.name == 'nt': # Проверяем, что мы на Windows
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE # Скрыть окно

                 self.backoffice_process = subprocess.Popen(
                    [backoffice_exe_path, backoffice_args],
                    cwd=installer_path,
                    shell=True, # Используем shell=True для корректной обработки аргументов
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0 # Не создавать консольное окно на Windows
                 )
                 log_message(f"BackOffice.exe успешно перезапущен (PID: {self.backoffice_process.pid}).")
                 self.update_status("BackOffice.exe успешно перезапущен.")
                 self.update_progress_step(1.0) # Шаг перезапуска завершен (достигает 100%)

                 self._launch_data['backoffice_process'] = self.backoffice_process # Обновляем объект процесса в данных запуска
                 # self.backoffice_process уже обновлен выше

            except FileNotFoundError:
                 raise FileNotFoundError(f"Не удалось найти исполняемый файл BackOffice.exe для перезапуска: '{backoffice_exe_path}'.")
            except Exception as e:
                 raise RuntimeError(f"Ошибка при перезапуске BackOffice.exe: {e}")


            # --- Завершение ---
            self.update_status("Готово! BackOffice запущен с обновленной конфигурацией.", level="INFO")
            self.update_progress(100) # Убеждаемся, что прогресс 100%
            log_message("Скрипт успешно завершен.", level="INFO")

        except Exception as e:
            self.handle_error(f"Произошла ошибка во время перезапуска BackOffice: {e}")

        finally:
            # Включаем ввод и кнопку обратно
            self.root.after(0, lambda: self.launch_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.target_entry.config(state=tk.NORMAL))


    def handle_error(self, message):
        """Обрабатывает ошибки: обновляет статус GUI, логирует, включает элементы GUI."""
        log_message(message, level="ERROR")
        # Обновление GUI в основном потоке
        self.root.after(0, lambda: self.status_label.config(text=f"Ошибка: {message}", foreground="red"))
        self.update_progress(0) # Сбрасываем прогресс при ошибке
        self.current_step_progress = 0.0
        self.progress_base = 0.0
        self.next_step_base = 0.0
        # Включаем элементы GUI обратно
        self.root.after(0, lambda: self.launch_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.target_entry.config(state=tk.NORMAL))

        # Сбрасываем объект процесса BackOffice, чтобы он не пытался остановиться при закрытии окна после ошибки
        # Проверяем наличие ключа, чтобы избежать ошибок, если _launch_data еще не полностью инициализирован
        if '_launch_data' in self.__dict__ and 'backoffice_process' in self._launch_data and self._launch_data['backoffice_process'] is not None:
            log_message(f"Попытка остановить процесс BackOffice PID {self._launch_data['backoffice_process'].pid} после ошибки...", level="WARNING")
            stop_process_by_pid(self._launch_data['backoffice_process'].pid)
            self._launch_data['backoffice_process'] = None # Сбрасываем в данных
        self.backoffice_process = None # Сбрасываем в классе


    def on_closing(self):
        """Обрабатывает закрытие окна - пытается остановить BackOffice, если он запущен."""
        log_message("Получен запрос на закрытие окна.")
        # Проверяем, запущен ли процесс BackOffice и не завершился ли он сам
        # Используем ссылку на процесс из _launch_data, если она есть
        current_process = None
        if '_launch_data' in self.__dict__ and 'backoffice_process' in self._launch_data:
             current_process = self._launch_data['backoffice_process']
        # else: # Запасной вариант, если _launch_data не инициализирован (например, при ошибке на ранних этапах)
        #      current_process = self.backoffice_process # Уже не нужно, т.к. self.backoffice_process дублирует _launch_data['backoffice_process']


        if current_process and current_process.poll() is None: # poll() возвращает код завершения или None, если процесс еще работает
            log_message(f"Обнаружен запущенный процесс BackOffice (PID: {current_process.pid}).")
            # Спрашиваем пользователя, нужно ли завершить процесс
            if messagebox.askokcancel("Выход", "BackOffice запущен. Завершить его перед выходом?"):
                log_message("Пользователь подтвердил завершение процесса BackOffice перед выходом.")
                self.update_status("Завершение процесса BackOffice...", level="INFO")
                stop_process_by_pid(current_process.pid)
                # Даем немного времени на завершение
                time.sleep(1.0)
                # Сбрасываем ссылки на процесс
                self.backoffice_process = None
                if '_launch_data' in self.__dict__ and 'backoffice_process' in self._launch_data:
                     self._launch_data['backoffice_process'] = None

                self.root.destroy() # Закрываем окно
            else:
                log_message("Пользователь отменил завершение процесса BackOffice. Окно останется открытым.")
                # Пользователь отменил, окно остается открытым
                pass
        else:
            log_message("Процесс BackOffice не запущен. Закрытие окна.")
            self.root.destroy() # Процесс не запущен, просто закрываем окно

# --- Основное выполнение ---
if __name__ == "__main__":

    # Проверяем аргументы командной строки
    initial_target = None
    if len(sys.argv) > 1:
        initial_target = sys.argv[1]
        # Логирование аргумента произойдет после загрузки конфига и установки global_debug_logging

    root = tk.Tk()
    app = BackOfficeLauncherGUI(root)

    # Если был аргумент командной строки, заполняем поле и запускаем процесс автоматически
    if initial_target:
        # Логируем аргумент после того, как global_debug_logging установлен
        log_message(f"Получен аргумент командной строки: '{initial_target}'", level="INFO")
        app.target_entry.insert(0, initial_target)
        # Запускаем процесс в отдельном потоке сразу после старта GUI
        # Используем root.after, чтобы дать GUI инициализироваться и отобразиться
        app.root.after(100, app.start_launch)


    root.protocol("WM_DELETE_WINDOW", app.on_closing) # Обрабатываем закрытие окна
    root.mainloop()
