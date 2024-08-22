#!/usr/bin/env python3

import argparse
import os
import pandas as pd
from ics import Calendar
import pytz
from typing import List, Dict, Optional
from datetime import datetime


def ics_to_excel(ics_file: str, excel_file: str, local_tz: str = 'Europe/Moscow') -> None:
    with open(ics_file, 'r', encoding='utf-8') as f:
        calendar = Calendar(f.read())

    events_data: List[Dict[str, Optional[str, datetime]]] = []

    local_timezone: datetime.tzinfo = pytz.timezone(local_tz)

    for event in calendar.events:
        # Перевод временных меток в локальный часовой пояс и удаление информации о часовом поясе
        local_start: datetime = event.begin.datetime.astimezone(local_timezone).replace(tzinfo=None)
        local_end: datetime = event.end.datetime.astimezone(local_timezone).replace(tzinfo=None)
        created: Optional[datetime] = event.created.datetime.astimezone(local_timezone).replace(
            tzinfo=None) if event.created else None
        last_modified: Optional[datetime] = event.last_modified.datetime.astimezone(local_timezone).replace(
            tzinfo=None) if event.last_modified else None

        events_data.append({
            'Название': event.name,
            'Начало': local_start,
            'Конец': local_end,
            'Описание': event.description,
            'Местоположение': event.location,
            'URL': event.url,
            'Создано': created,
            'Последнее изменение': last_modified,
            'Статус': event.status,
            'Организатор': event.organizer.common_name if event.organizer else None,
            'Участники': ', '.join([attendee.common_name for attendee in event.attendees]) if event.attendees else None,
            'Категории': ', '.join(event.categories) if event.categories else None,
            'Прозрачность': event.transparent
        })

    df = pd.DataFrame(events_data)
    df.to_excel(excel_file, index=False)

    print(f"Данные успешно сохранены в {excel_file}")


def main():
    parser = argparse.ArgumentParser(description='Конвертирование ICS файла в Excel файл.')
    parser.add_argument(
        '-i', '--input',
        help='Имя входного файла ICS. По умолчанию будет произведен поиск всех .ics файлов в текущей директории.'
    )
    parser.add_argument(
        '-o', '--output',
        help='Имя выходного файла Excel. По умолчанию - имя входного файла, но с расширением .xlsx'
    )
    parser.add_argument(
        '-t', '--timezone',
        help='Часовой пояс. По умолчанию - "Europe/Moscow"',
        default='Europe/Moscow'
    )

    args = parser.parse_args()

    if args.input:
        ics_file: str = args.input
        excel_file: str = args.output if args.output else os.path.splitext(ics_file)[0] + '.xlsx'
        ics_to_excel(ics_file, excel_file, args.timezone)
    else:
        # Поиск всех файлов .ics в текущей директории
        current_directory = os.getcwd()
        ics_files = [f for f in os.listdir(current_directory) if f.endswith('.ics')]

        if not ics_files:
            print("В текущей директории нет файлов с расширением .ics")
        else:
            for ics_file in ics_files:
                excel_file: str = os.path.splitext(ics_file)[0] + '.xlsx'  # Убираем расширение .ics и добавляем .xlsx
                ics_to_excel(ics_file, excel_file, args.timezone)


if __name__ == '__main__':
    main()
