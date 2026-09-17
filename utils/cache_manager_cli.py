#!/usr/bin/env python3
"""
Утилита командной строки для управления кэшем DS данных
"""

import sys
import argparse
from utils.cache_utils import cache_manager


def main():
    parser = argparse.ArgumentParser(description='Управление кэшем DS данных')
    parser.add_argument('action', choices=['info', 'clear', 'clear-file'], 
                       help='Действие: info - показать информацию о кэше, clear - очистить весь кэш, clear-file - очистить кэш для конкретного файла')
    parser.add_argument('--file', '-f', help='Путь к файлу для очистки кэша (только для действия clear-file)')
    
    args = parser.parse_args()
    
    if args.action == 'info':
        info = cache_manager.get_cache_info()
        print("Информация о кэше:")
        print(f"  Директория кэша: {info['cache_dir']}")
        print(f"  Количество файлов: {info['files_count']}")
        print(f"  Общий размер: {info['total_size_mb']} MB")
        if info['files']:
            print("  Файлы кэша:")
            for file_path in info['files']:
                print(f"    - {file_path}")
        else:
            print("  Кэш пуст")
    
    elif args.action == 'clear':
        if cache_manager.clear_cache():
            print("Кэш успешно очищен")
        else:
            print("Ошибка при очистке кэша")
            sys.exit(1)
    
    elif args.action == 'clear-file':
        if not args.file:
            print("Ошибка: для действия clear-file необходимо указать файл с помощью --file")
            sys.exit(1)
        
        if cache_manager.clear_cache(args.file):
            print(f"Кэш для файла {args.file} успешно очищен")
        else:
            print(f"Ошибка при очистке кэша для файла {args.file}")
            sys.exit(1)


if __name__ == '__main__':
    main()
