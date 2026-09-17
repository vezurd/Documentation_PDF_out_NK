from typing import Dict, List, Optional, Any, Union
from enum import Enum
from dataclasses import dataclass
import utils.path
from base.base_class_std_table import STDTable
from base.base_classes import *
import base.t_comm_initial_classes
import base.base_mto
import RFQ.ds_compare.get_mto_list_from_ds
import RFQ.ds_compare.load_mto
from base.tables_columns import *
from base.base_classes import RowStd, RowType
from base.base_excel_out import *
from base.tables_columns import ColNames
from utils.string_parsing import print_att_list_table
from RFQ.ds_compare.ds_classes import *

def generate_report(results: Dict[str, Any], ds_path: str) -> None:
    """
    Генерирует различные отчеты на основе результатов сравнения
    """
    report_dir = utils.path.get_path_from_file_path(ds_path)

    # 1. Детальный отчет в Excel
    generate_detailed_excel_report(results, report_dir)

    # 2. Сводный отчет в Excel
    generate_summary_excel_report(results, report_dir)

    # 3. Текстовый отчет
    generate_text_report(results, report_dir)

    # 4. Отчет по проблемным позициям
    generate_issues_report(results, report_dir)

    # 5. Лог выполнения
    generate_execution_log(results, report_dir)


def generate_detailed_excel_report(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует детальный Excel отчет"""
    try:
        import pandas as pd
        from datetime import datetime

        # Создаем DataFrame с результатами
        report_data = []
        for result in results.get('comparison_results', []):
            report_data.append({
                'Спецификация': result.get('spec_name', ''),
                'Код': result.get('code', ''),
                'Наименование': result.get('name', ''),
                'Количество DS': result.get('ds_amount', 0),
                'Количество MTO': result.get('mto_amount', 0),
                'Разница': result.get('difference', 0),
                'Статус': result.get('status', ''),
                'Комментарий': result.get('comment', '')
            })

        if report_data:
            df = pd.DataFrame(report_data)

            # Добавляем стили
            def color_status(val):
                if val == 'Совпадает':
                    return 'background-color: #C6EFCE'  # Зеленый
                elif val == 'Недостаточно':
                    return 'background-color: #FFC7CE'  # Красный
                elif val == 'Избыточно':
                    return 'background-color: #FFEB9C'  # Желтый
                elif val == 'Не найдено':
                    return 'background-color: #FFCC99'  # Оранжевый
                return ''

            styled_df = df.style.applymap(color_status, subset=['Статус'])

            # Сохраняем отчет
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = f"{report_dir}/Детальный_отчет_сравнения_{timestamp}.xlsx"
            styled_df.to_excel(report_path, index=False)
            print(f"Детальный отчет сохранен: {report_path}")

    except ImportError:
        print("Для генерации Excel отчетов требуется установить pandas: pip install pandas")


def generate_summary_excel_report(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует сводный Excel отчет"""
    try:
        import pandas as pd
        from datetime import datetime

        stats = results.get('statistics', {})
        summary_data = {
            'Показатель': [
                'Всего строк в DS',
                'Совпадает количеств',
                'Недостаточно в MTO',
                'Избыточно в MTO',
                'Не найдено в MTO',
                'Новых позиций добавлено',
                'Процент совпадения'
            ],
            'Значение': [
                stats.get('total_ds_rows', 0),
                stats.get('matched', 0),
                stats.get('insufficient', 0),
                stats.get('excess', 0),
                stats.get('not_found', 0),
                stats.get('new_positions_count', 0),
                f"{(stats.get('matched', 0) / stats.get('total_ds_rows', 1) * 100):.1f}%"
            ]
        }

        df = pd.DataFrame(summary_data)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = f"{report_dir}/Сводный_отчет_{timestamp}.xlsx"
        df.to_excel(report_path, index=False)
        print(f"Сводный отчет сохранен: {report_path}")

    except ImportError:
        print("Для генерации Excel отчетов требуется установить pandas")


def generate_text_report(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует текстовый отчет"""
    from datetime import datetime

    stats = results.get('statistics', {})
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report_content = f"""
ОТЧЕТ СРАВНЕНИЯ DS И MTO
Дата генерации: {timestamp}
========================================

СТАТИСТИКА:
-----------
Всего строк в DS: {stats.get('total_ds_rows', 0)}
Совпадает количеств: {stats.get('matched', 0)}
Недостаточно в MTO: {stats.get('insufficient', 0)}
Избыточно в MTO: {stats.get('excess', 0)}
Не найдено в MTO: {stats.get('not_found', 0)}
Новых позиций добавлено: {stats.get('new_positions_count', 0)}
Процент совпадения: {(stats.get('matched', 0) / max(stats.get('total_ds_rows', 1), 1) * 100):.1f}%

НОВЫЕ ПОЗИЦИИ:
--------------
"""

    new_positions = results.get('new_positions', [])
    if new_positions:
        for i, pos in enumerate(new_positions, 1):
            report_content += f"{i}. {pos.get('spec_name', '')}: {pos.get('code', '')} - {pos.get('name', '')} ({pos.get('amount', 0)})\n"
    else:
        report_content += "Новых позиций не обнаружено\n"

    report_content += """
ПРОБЛЕМНЫЕ ПОЗИЦИИ:
------------------
"""

    # Добавляем информацию о проблемных позициях
    problem_codes = set()
    for result in results.get('comparison_results', []):
        status = result.get('status', '')
        if status in ['Недостаточно', 'Не найдено']:
            problem_codes.add((result.get('spec_name', ''), result.get('code', '')))

    if problem_codes:
        for i, (spec, code) in enumerate(problem_codes, 1):
            report_content += f"{i}. {spec}: {code}\n"
    else:
        report_content += "Проблемных позиций не обнаружено\n"

    # Сохраняем отчет
    report_path = f"{report_dir}/Текстовый_отчет_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report_content)

    print(f"Текстовый отчет сохранен: {report_path}")


def generate_issues_report(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует отчет по проблемным позициям"""
    from datetime import datetime

    issues = []
    for result in results.get('comparison_results', []):
        status = result.get('status', '')
        if status in ['Недостаточно', 'Не найдено']:
            issues.append({
                'spec_name': result.get('spec_name', ''),
                'code': result.get('code', ''),
                'name': result.get('name', ''),
                'ds_amount': result.get('ds_amount', 0),
                'mto_amount': result.get('mto_amount', 0),
                'status': status,
                'comment': result.get('comment', '')
            })

    if issues:
        report_content = "ОТЧЕТ ПО ПРОБЛЕМНЫМ ПОЗИЦИЯМ\n"
        report_content += "=" * 50 + "\n\n"

        for i, issue in enumerate(issues, 1):
            report_content += f"{i}. Спецификация: {issue['spec_name']}\n"
            report_content += f"   Код: {issue['code']}\n"
            report_content += f"   Наименование: {issue['name']}\n"
            report_content += f"   Количество DS: {issue['ds_amount']}\n"
            report_content += f"   Количество MTO: {issue['mto_amount']}\n"
            report_content += f"   Статус: {issue['status']}\n"
            report_content += f"   Комментарий: {issue['comment']}\n"
            report_content += "-" * 50 + "\n"

        report_path = f"{report_dir}/Проблемные_позиции_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)

        print(f"Отчет по проблемным позициям сохранен: {report_path}")


def generate_execution_log(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует лог выполнения"""
    from datetime import datetime

    log_content = f"""
ЛОГ ВЫПОЛНЕНИЯ СРАВНЕНИЯ
Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
========================================

ОБЩАЯ ИНФОРМАЦИЯ:
-----------------
Время начала: {results.get('start_time', 'N/A')}
Время завершения: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Обработано спецификаций: {len(results.get('processed_specs', []))}

ОБРАБОТАННЫЕ СПЕЦИФИКАЦИИ:
--------------------------
"""

    processed_specs = results.get('processed_specs', [])
    for spec in processed_specs:
        log_content += f"- {spec}\n"

    log_content += f"""
СТАТИСТИКА ОБРАБОТКИ:
--------------------
Всего строк: {results.get('statistics', {}).get('total_ds_rows', 0)}
Успешно: {results.get('statistics', {}).get('matched', 0)}
С предупреждениями: {results.get('statistics', {}).get('excess', 0)}
С ошибками: {results.get('statistics', {}).get('insufficient', 0) + results.get('statistics', {}).get('not_found', 0)}

ОШИБКИ И ПРЕДУПРЕЖДЕНИЯ:
------------------------
"""

    errors = results.get('errors', [])
    if errors:
        for error in errors:
            log_content += f"ERROR: {error}\n"
    else:
        log_content += "Ошибок не обнаружено\n"

    warnings = results.get('warnings', [])
    if warnings:
        for warning in warnings:
            log_content += f"WARNING: {warning}\n"
    else:
        log_content += "Предупреждений не обнаружено\n"

    # Сохраняем лог
    log_path = f"{report_dir}/Лог_выполнения_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(log_content)

    print(f"Лог выполнения сохранен: {log_path}")


def generate_html_report(results: Dict[str, Any], report_dir: str) -> None:
    """Генерирует HTML отчет (опционально)"""
    try:
        from datetime import datetime
        import json

        # Создаем данные для HTML отчета
        report_data = {
            'generated_at': datetime.now().isoformat(),
            'statistics': results.get('statistics', {}),
            'new_positions': results.get('new_positions', []),
            'issues': [
                result for result in results.get('comparison_results', [])
                if result.get('status') in ['Недостаточно', 'Не найдено']
            ]
        }

        # Сохраняем данные в JSON для использования в HTML
        json_path = f"{report_dir}/report_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, ensure_ascii=False, indent=2)

        print(f"Данные для HTML отчета сохранены: {json_path}")

    except Exception as e:
        print(f"Ошибка при генерации HTML отчета: {e}")


def open_report_directory(report_dir: str) -> None:
    """Открывает папку с отчетами в проводнике"""
    import os
    import subprocess

    try:
        if os.name == 'nt':  # Windows
            os.startfile(report_dir)
        elif os.name == 'posix':  # macOS, Linux
            subprocess.run(['open', report_dir] if os.uname().sysname == 'Darwin' else ['xdg-open', report_dir])
        print(f"Открыта папка с отчетами: {report_dir}")
    except Exception as e:
        print(f"Не удалось открыть папку: {e}")


def cleanup_old_reports(report_dir: str, days_to_keep: int = 30) -> None:
    """Очищает старые отчеты"""
    import os
    import time
    from datetime import datetime, timedelta

    try:
        cutoff_time = time.time() - (days_to_keep * 24 * 60 * 60)

        for filename in os.listdir(report_dir):
            if filename.startswith(('Детальный_отчет', 'Сводный_отчет', 'Текстовый_отчет',
                                    'Проблемные_позиции', 'Лог_выполнения')):
                file_path = os.path.join(report_dir, filename)
                if os.path.getmtime(file_path) < cutoff_time:
                    os.remove(file_path)
                    print(f"Удален старый отчет: {filename}")

    except Exception as e:
        print(f"Ошибка при очистке старых отчетов: {e}")