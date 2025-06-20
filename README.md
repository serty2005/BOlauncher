# BackOffice Launcher

Приложение для автоматизированного запуска iiko/Syrve BackOffice с настройкой конфигурации и загрузкой дистрибутивов.

## Функционал

### Основные возможности
- **Автоматическая загрузка дистрибутивов** с нескольких источников:
  - HTTP/HTTPS сервер
  - FTP сервер
  - SMB-шары (локальные/сетевые)
- **Поддержка двух производителей**:
  - `iiko` (RMS и Chain)
  - `Syrve` (RMS и Chain)
- **Умное определение версии** из строки сервера
- **Автоматическая настройка конфига**:
  - Ожидание создания конфиг-файла
  - Парсинг и модификация XML
- **Очистка кэша AppData** перед запуском
- **Контроль процессов**:
  - Остановка работающего BackOffice перед редактированием конфига
  - Перезапуск после настройки
- **Графический интерфейс** с прогресс-баром и логом операций

### Дополнительные функции
- Проверка сервера через REST API (`/resto/getServerMonitoringInfo.jsp`)
- Валидация производителя по метаданным EXE-файла
- Обработка ошибок с детализированным логом
- Поддержка аргументов командной строки

Вы можете запустить скрипт с адресом в качестве аргумента. В этом случае GUI откроется, поле ввода будет заполнено аргументом, и процесс запуска начнется автоматически.

### Особенности конфигурации

  Приоритет источников - проверяются в порядке, указанном в Order
  Подстановки в именах архивов:
    {version} заменяется на форматированную версию (например "887" для версии 8.8.7)

  Безопасность:
    Пароли FTP хранятся в открытом виде!
    Для SMB используйте доступ только для чтения

## Конфигурация (`config.ini`)

Файл создается автоматически при первом запуске. Пример структуры:

```ini
[Settings]
; Таймаут HTTP-запросов (сек)
HttpRequestTimeoutSec = 15

; Корневая папка для локальных дистрибутивов
InstallerRoot = C:\iiko_Distr

; Таймаут ожидания конфиг-файла (сек)
ConfigFileWaitTimeoutSec = 60

; Интервал проверки конфиг-файла (мс)
ConfigFileCheckIntervalMs = 100

; Включить подробное логирование (True/False)
DebugLogging = False

[SourcePriority]
; Порядок проверки источников (через запятую)
Order = smb, http, ftp

[SmbSource]
; Включить SMB источник
Enabled = False

; UNC-путь к папке с архивами
Path = \\10.25.100.5\sharedisk\iikoBacks

; Шаблоны имен архивов (подстановка {version} и {vendor_subdir})
iikoRMS_ArchiveName = RMSOffice{version}.zip
iikoChain_ArchiveName = ChainOffice{version}.zip
SyrveRMS_ArchiveName = Syrve/RMSSOffice{version}.zip
SyrveChain_ArchiveName = Syrve/ChainSOffice{version}.zip

[HttpSource]
; Включить HTTP источник
Enabled = True

; Базовый URL
Url = https://f.serty.top/iikoBacks

; Шаблоны имен архивов
iikoRMS_ArchiveName = RMSOffice{version}.zip
iikoChain_ArchiveName = ChainOffice{version}.zip
SyrveRMS_ArchiveName = Syrve/RMSSOffice{version}.zip
SyrveChain_ArchiveName = Syrve/ChainSOffice{version}.zip

[FtpSource]
; Включить FTP источник
Enabled = True

; Параметры подключения
Host = ftp.serty.top
Port = 21
Username = ftpuser
Password = 22  # Внимание: пароль хранится в открытом виде
Directory = /iikoBacks

; Шаблоны имен архивов
iikoRMS_ArchiveName = RMSOffice{version}.zip
iikoChain_ArchiveName = ChainOffice{version}.zip
SyrveRMS_ArchiveName = Syrve/RMSSOffice{version}.zip
SyrveChain_ArchiveName = Syrve/ChainSOffice{version}.zip

[LocalInstallerNames]
; Форматы имен локальных папок
iikoRMS = RMSOffice
iikoChain = ChainOffice
SyrveRMS = RMSSOffice
SyrveChain = ChainSOffice

```

