# Windows Server Network Adapter Report

Русская редакция PowerShell-скрипта для сбора сведений о сетевых адаптерах Windows Server, default gateway,
сохранённых TCP/IP-настройках и старых сетевых профилях из реестра.

## Для чего нужен скрипт

Скрипт помогает быстро получить единый отчёт по списку серверов, когда нужно:

- посчитать сетевые адаптеры на каждом сервере
- собрать IP, DNS, DHCP, MAC, скорость и статус подключений
- увидеть прописанные шлюзы в конфигурации адаптеров и в таблице маршрутов
- отдельно найти серверы с несколькими default gateway
- выгрузить сведения о старых или уже удалённых адаптерах, у которых в системе остались TCP/IP-профили

## Как работает скрипт

Скрипт читает список серверов из `servers.txt` и подключается к ним через WMI/DCOM и удалённый реестр
(`StdRegProv`). Такой подход выбран специально для совместимости с Windows Server 2008-2025 и не зависит от
современных модулей `NetAdapter` и `NetTCPIP`.

Логика работы:

1. Читает TXT-файл со списком серверов.
2. Предлагает использовать текущую учётную запись или запросить другие учётные данные.
3. Подключается к каждому серверу и собирает сведения о текущих адаптерах, IP-конфигурации и default gateway.
4. Читает ветки реестра с TCP/IP-профилями и class registry, чтобы выявить старые сетевые профили.
5. Исключает служебные адаптеры `WAN Miniport` и `Microsoft`.
6. Проверяет, есть ли у сервера несколько шлюзов по умолчанию, и выносит такие случаи в отдельный отчёт.
7. Сохраняет результат в CSV и JSON.

## Основные возможности

- поддержка Windows Server 2008, 2008 R2, 2012, 2012 R2, 2016, 2019, 2022 и 2025
- работа из Windows PowerShell 5.1 и PowerShell 7+
- интерактивный выбор между текущим пользователем и `Get-Credential`
- многопоточная обработка списка серверов через фоновые jobs
- автоматический запуск worker-процессов Windows PowerShell 5.1 для WMI-сбора, если основной запуск идёт из
  PowerShell 7+
- отдельный отчёт по серверам с несколькими default gateway
- более читаемый отчёт по старым адаптерам и сохранённым настройкам

## Состав русской редакции

- [get_windows_server_network_adapter_report.ps1](./get_windows_server_network_adapter_report.ps1) — основной скрипт
- [servers.txt.example](./servers.txt.example) — пример файла со списком серверов

## Требования

- доступ к удалённым серверам по WMI/DCOM
- права на чтение WMI и удалённого реестра
- на машине запуска должен быть доступен Windows PowerShell 5.1 для WMI-worker-процессов
- открытые RPC/WMI-порты между машиной запуска и целевыми серверами

## Подготовка списка серверов

Создайте `servers.txt` рядом со скриптом на основе `servers.txt.example`.

Пример:

```text
SRV-DC-01
SRV-FS-01
# строки с # игнорируются
SRV-APP-01
```

## Примеры запуска

```powershell
.\get_windows_server_network_adapter_report.ps1
.\get_windows_server_network_adapter_report.ps1 -UseCurrentUser
.\get_windows_server_network_adapter_report.ps1 -ComputerListPath .\servers.txt
.\get_windows_server_network_adapter_report.ps1 -ComputerListPath .\servers.txt -Credential (Get-Credential)
.\get_windows_server_network_adapter_report.ps1 -Parallel
.\get_windows_server_network_adapter_report.ps1 -Parallel -ThrottleLimit 12
.\get_windows_server_network_adapter_report.ps1 -OutputDirectory .\output\manual_run
```

## Параметры

- `-ComputerListPath` — путь к TXT-файлу со списком серверов
- `-OutputDirectory` — папка для выгрузки
- `-Credential` — явные учётные данные для WMI/DCOM
- `-UseCurrentUser` — использовать текущего пользователя без запроса учётных данных
- `-Parallel` — включить многопоточную обработку
- `-ThrottleLimit` — ограничить число одновременно работающих jobs

## Какие файлы создаются

По умолчанию выгрузка сохраняется в `.\output\<yyyyMMdd_HHmmss>`.

- `network_adapter_summary.csv` — краткая сводка по каждому серверу
- `network_adapter_details.csv` — текущие сетевые адаптеры и их настройки
- `network_gateway_details.csv` — шлюзы из конфигурации адаптеров и из route table
- `network_multiple_gateways.csv` — серверы, где найдено несколько default gateway
- `network_legacy_adapters.csv` — человекочитаемый список старых адаптеров и сохранённых настроек
- `network_adapter_report.json` — полный структурированный отчёт

## Как читать отчёт по старым адаптерам

В `network_legacy_adapters.csv` основные колонки такие:

- `LegacyCategory` — тип найденного следа старого адаптера
- `WhatWasFound` — краткое пояснение, что осталось в системе
- `AdapterName` — понятное имя адаптера или драйвера
- `AddressingMode` — сохранённый тип адресации
- `DHCPEnabled` — был ли включён DHCP
- `SavedIPAddress` — сохранённые IP-адреса
- `SavedSubnetMask` — сохранённые маски
- `SavedDefaultGateway` — сохранённые шлюзы
- `SavedDnsServers` — сохранённые DNS-серверы
- `SavedSettingsSummary` — короткая сводка по найденным параметрам

## Примечания

- скрипт не требует PowerShell Remoting или WinRM
- закрытый ICMP не мешает сбору, если доступен WMI/DCOM
- на старых ОС нет надёжного универсального признака скрытого адаптера, поэтому используются категории
  `LegacyWithTcpipProfile`, `RegistryOnly` и `ClassOnly`
- при большом списке серверов подбирайте `ThrottleLimit` под нагрузку на сеть и машину запуска

## Лицензия

Проект распространяется по лицензии MIT. Корневой файл лицензии: [../LICENSE](../LICENSE)
