import os
import pickle
import hashlib
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Optional, Tuple, Dict, Callable
from pathlib import Path
import base.base_classes
from base.base_classes import RowStd


class CacheManager:
    """Менеджер кэша для данных DS файлов"""
    
    def __init__(self, cache_dir: str = "cache"):
        """
        Инициализация менеджера кэша
        
        Args:
            cache_dir: Директория для хранения кэш-файлов
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(exist_ok=True)
    
    def _get_path_hash(self, file_path: str) -> str:
        """
        Вычисляет MD5 хэш пути к файлу (без чтения файла).
        Используется для именования кэш-файлов.
        """
        normalized = os.path.normpath(os.path.abspath(file_path))
        return hashlib.md5(normalized.encode(errors='replace')).hexdigest()[:16]

    def _get_cache_file_path(self, file_path: str, extra_key: Optional[str] = None) -> Path:
        """
        Получает путь к кэш-файлу на основе пути к исходному файлу.
        При extra_key добавляет его хэш в имя для инвалидации при смене конфига.
        """
        path_hash = self._get_path_hash(file_path)
        if extra_key:
            extra_hash = hashlib.md5(extra_key.encode(errors='replace')).hexdigest()[:12]
            return self.cache_dir / f"v2_{path_hash}_{extra_hash}.cache"
        return self.cache_dir / f"v2_{path_hash}.cache"
    
    def is_cache_valid(self, file_path: str, extra_key: Optional[str] = None) -> bool:
        """
        Проверяет, действителен ли кэш для указанного файла.
        Использует mtime (stat) — без чтения исходного файла.
        """
        if not os.path.exists(file_path):
            return False
        cache_file_path = self._get_cache_file_path(file_path, extra_key)
        if not cache_file_path.exists():
            return False
        try:
            file_mtime = os.path.getmtime(file_path)
            with open(cache_file_path, 'rb') as f:
                cached = pickle.load(f)
            stored_mtime = cached.get('mtime') if isinstance(cached, dict) else None
            return stored_mtime is not None and abs(stored_mtime - file_mtime) < 1e-6
        except Exception:
            return False
    
    def save_to_cache(self, file_path: str, data: List[RowStd], verbose: bool = True, extra_key: Optional[str] = None) -> bool:
        """
        Сохраняет данные в кэш с mtime для последующей проверки валидности.
        """
        try:
            cache_file_path = self._get_cache_file_path(file_path, extra_key)
            file_mtime = os.path.getmtime(file_path) if os.path.exists(file_path) else 0.0
            cached = {'mtime': file_mtime, 'data': data}
            with open(cache_file_path, 'wb') as f:
                pickle.dump(cached, f)
            if verbose:
                print(f"    Данные сохранены в кэш: {cache_file_path}")
            return True
        except Exception as e:
            print(f"Ошибка при сохранении в кэш: {e}")
            return False
    
    def load_from_cache(self, file_path: str, verbose: bool = True, extra_key: Optional[str] = None) -> Optional[List[RowStd]]:
        """
        Загружает данные из кэша. Проверяет mtime (без чтения исходного файла).
        """
        try:
            cache_file_path = self._get_cache_file_path(file_path, extra_key)
            if not cache_file_path.exists():
                return None
            with open(cache_file_path, 'rb') as f:
                cached = pickle.load(f)
            if isinstance(cached, dict) and 'data' in cached:
                file_mtime = os.path.getmtime(file_path) if os.path.exists(file_path) else 0.0
                stored_mtime = cached.get('mtime')
                if stored_mtime is not None and abs(stored_mtime - file_mtime) < 1e-6:
                    if verbose:
                        print(f"    Данные загружены из кэша: {cache_file_path}")
                    return cached['data']
            return None
        except Exception as e:
            print(f"Ошибка при загрузке из кэша: {e}")
            return None
    
    def clear_cache(self, file_path: str = None, extra_key: Optional[str] = None) -> bool:
        """
        Очищает кэш для указанного файла или весь кэш
        
        Args:
            file_path: Путь к файлу для очистки кэша (если None - очищает весь кэш)
            extra_key: Дополнительный ключ (если использовался при сохранении)
            
        Returns:
            True если очистка успешна, False иначе
        """
        try:
            if file_path:
                cache_file_path = self._get_cache_file_path(file_path, extra_key)
                if cache_file_path and cache_file_path.exists():
                    cache_file_path.unlink()
                    print(f"Кэш для файла {file_path} очищен")
            else:
                # Очищаем весь кэш
                for cache_file in self.cache_dir.glob("*.cache"):
                    cache_file.unlink()
                print("Весь кэш очищен")
            return True
        except Exception as e:
            print(f"Ошибка при очистке кэша: {e}")
            return False
    
    def _get_mto_agg_cache_path(self, mto_path: str, extra_key: Optional[str] = None) -> Path:
        """Путь к кэш-файлу агрегированных mto_data для директории."""
        path_hash = hashlib.md5(mto_path.encode(errors='replace')).hexdigest()[:16]
        if extra_key:
            extra_hash = hashlib.md5(extra_key.encode(errors='replace')).hexdigest()[:12]
            return self.cache_dir / f"mto_agg_{path_hash}_{extra_hash}.cache"
        return self.cache_dir / f"mto_agg_{path_hash}.cache"

    def load_mto_data_from_cache(
        self,
        mto_path: str,
        verbose: bool = False,
        extra_key: Optional[str] = None
    ) -> Optional[Dict[str, Dict]]:
        """
        Загружает кэш агрегации MTO. Возвращает file_contributions:
        {file_path: {'mtime': float, 'title_system': str, 'rows': List[RowStd]}}
        или None.
        """
        cache_path = self._get_mto_agg_cache_path(mto_path, extra_key)
        if not cache_path.exists():
            return None
        try:
            with open(cache_path, 'rb') as f:
                data = pickle.load(f)
            contributions = data.get('file_contributions') if isinstance(data, dict) else None
            if contributions is None:
                return None
            if verbose:
                print(f"    mto_agg кэш загружен: {cache_path}")
            return contributions
        except Exception as e:
            print(f"Ошибка при загрузке mto_agg кэша: {e}")
            return None

    def save_mto_data_to_cache(
        self,
        mto_path: str,
        file_contributions: Dict[str, Dict],
        verbose: bool = False,
        extra_key: Optional[str] = None
    ) -> bool:
        """Сохраняет агрегированный mto_agg кэш."""
        try:
            cache_path = self._get_mto_agg_cache_path(mto_path, extra_key)
            data = {'file_contributions': file_contributions}
            with open(cache_path, 'wb') as f:
                pickle.dump(data, f)
            if verbose:
                print(f"    mto_agg кэш сохранён: {cache_path}")
            return True
        except Exception as e:
            print(f"Ошибка при сохранении mto_agg кэша: {e}")
            return False

    # Версия кэша VO: при исправлении ошибок загрузки (ANNOTATION и др.) — инвалидирует старый кэш
    VO_AGG_CACHE_VERSION = "v2"  # v2: исправлен NameError ANNOTATION в VO_MTO

    def _get_vo_agg_path_hash(self, vo_path: str) -> str:
        """Хэш пути для vo_agg кэша (включает версию для инвалидации при изменении кода)."""
        key = f"{vo_path}|{self.VO_AGG_CACHE_VERSION}"
        return hashlib.md5(key.encode(errors='replace')).hexdigest()[:16]

    def _get_vo_agg_cache_path(self, vo_path: str) -> Path:
        """Путь к кэш-файлу агрегированных vo_data (старый формат, для обратной совместимости)."""
        path_hash = self._get_vo_agg_path_hash(vo_path)
        return self.cache_dir / f"vo_agg_{path_hash}.cache"

    def _get_vo_agg_meta_path(self, vo_path: str) -> Path:
        """Путь к мета-файлу vo_agg (малый, для быстрой проверки валидности)."""
        path_hash = self._get_vo_agg_path_hash(vo_path)
        return self.cache_dir / f"vo_agg_meta_{path_hash}.cache"

    def _get_vo_agg_data_path(self, vo_path: str) -> Path:
        """Путь к файлу данных vo_agg (большой pickle)."""
        path_hash = self._get_vo_agg_path_hash(vo_path)
        return self.cache_dir / f"vo_agg_data_{path_hash}.cache"

    def _is_vo_data_cache_valid(
        self,
        cache_mtime: float,
        file_paths: List[str],
        cached_paths: Optional[List[str]],
        max_workers: int = 16
    ) -> bool:
        """
        Проверяет валидность кэша: тот же набор файлов и ни один не менялся после создания кэша.
        Проверка mtime выполняется параллельно (важно для сетевых путей).
        """
        if cached_paths is None:
            return False
        if sorted(cached_paths) != sorted(file_paths):
            return False
        try:
            def _check_file(fp: str) -> bool:
                """True если файл не менялся (mtime <= cache_mtime)."""
                try:
                    if not os.path.exists(fp):
                        return True
                    return os.path.getmtime(fp) <= cache_mtime
                except Exception:
                    return False

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(_check_file, fp): fp for fp in file_paths}
                for future in as_completed(futures):
                    if not future.result():
                        return False
            return True
        except Exception:
            return False

    def save_vo_data_to_cache(
        self,
        vo_path: str,
        vo_data: Dict[str, List[RowStd]],
        empty_title_rows: Dict[str, List[RowStd]],
        file_paths: List[str],
        verbose: bool = False
    ) -> bool:
        """Сохраняет агрегированные vo_data в кэш (мета + данные отдельно для быстрой загрузки)."""
        try:
            meta_path = self._get_vo_agg_meta_path(vo_path)
            data_path = self._get_vo_agg_data_path(vo_path)
            cache_mtime = time.time()

            meta = {'file_paths': file_paths, 'cache_mtime': cache_mtime}
            with open(meta_path, 'wb') as f:
                pickle.dump(meta, f)

            data = {
                'vo_data': dict(vo_data),
                'empty_title_rows': dict(empty_title_rows),
            }
            with open(data_path, 'wb') as f:
                pickle.dump(data, f)

            if verbose:
                print(f"    vo_data сохранены в кэш: {data_path}")
            return True
        except Exception as e:
            print(f"Ошибка при сохранении vo_data в кэш: {e}")
            return False

    def load_vo_data_from_cache(
        self,
        vo_path: str,
        file_paths: List[str],
        verbose: bool = False,
        result_dir: Optional[str] = None,
        log_timing: Optional[Callable[[str, str], None]] = None
    ) -> Optional[Tuple[Dict[str, List[RowStd]], Dict[str, List[RowStd]]]]:
        """
        Загружает агрегированные vo_data из кэша.
        Сначала проверяет мета-файл (малый), затем загружает данные только при валидном кэше.
        Поддерживает старый формат (один файл) для обратной совместимости.
        При переданных result_dir и log_timing пишет подробный лог времени выполнения.
        """
        def _log(msg: str, t: float) -> None:
            if result_dir and log_timing:
                log_timing(result_dir, f"step3_vo_cache_{msg}: {t:.3f}s")

        meta_path = self._get_vo_agg_meta_path(vo_path)
        data_path = self._get_vo_agg_data_path(vo_path)
        legacy_path = self._get_vo_agg_cache_path(vo_path)

        try:
            if meta_path.exists() and data_path.exists():
                t0 = time.perf_counter()
                with open(meta_path, 'rb') as f:
                    meta = pickle.load(f)
                _log("meta_load", time.perf_counter() - t0)

                cached_paths = meta.get('file_paths') if isinstance(meta, dict) else None
                cache_mtime = meta.get('cache_mtime', 0.0)

                t0 = time.perf_counter()
                valid = self._is_vo_data_cache_valid(cache_mtime, file_paths, cached_paths)
                _log("mtime_validation", time.perf_counter() - t0)

                if not valid:
                    _log("cache_miss_invalid", 0.0)
                    return None

                t0 = time.perf_counter()
                with open(data_path, 'rb') as f:
                    data = pickle.load(f)
                _log("data_load", time.perf_counter() - t0)

                vo_data = data.get('vo_data', {})
                empty_title_rows = data.get('empty_title_rows', {})
                if verbose:
                    print(f"    vo_data загружены из кэша: {data_path}")
                return (vo_data, empty_title_rows)

            if legacy_path.exists():
                t0 = time.perf_counter()
                with open(legacy_path, 'rb') as f:
                    data = pickle.load(f)
                _log("legacy_full_load", time.perf_counter() - t0)

                cached_paths = data.get('file_paths') if isinstance(data, dict) else None

                t0 = time.perf_counter()
                valid = self._is_vo_data_cache_valid(
                    os.path.getmtime(legacy_path), file_paths, cached_paths
                )
                _log("legacy_mtime_validation", time.perf_counter() - t0)

                if not valid:
                    _log("legacy_cache_miss_invalid", 0.0)
                    return None

                vo_data = data.get('vo_data', {})
                empty_title_rows = data.get('empty_title_rows', {})
                if verbose:
                    print(f"    vo_data загружены из кэша (legacy): {legacy_path}")
                return (vo_data, empty_title_rows)

            _log("cache_miss_no_files", 0.0)
            return None
        except Exception as e:
            print(f"Ошибка при загрузке vo_data из кэша: {e}")
            return None

    # Версия кэша step1: при изменении логики get_row_copy_light (и др.) — инвалидирует старый кэш
    RFP_STEP1_CACHE_VERSION = "v3"  # v3: RFP_AGGREGATED VALUES_2 optional (net without col R)

    def _get_rfp_step1_cache_path(self, rfp_path: str, code_ban_file: Optional[str], extra_key: Optional[str] = None) -> Path:
        """Путь к кэш-файлу полного результата step1_load_rfp."""
        key = f"{rfp_path}|{code_ban_file or ''}|{self.RFP_STEP1_CACHE_VERSION}"
        if extra_key:
            key = f"{key}|{extra_key}"
        path_hash = hashlib.md5(key.encode(errors='replace')).hexdigest()[:16]
        return self.cache_dir / f"rfp_step1_{path_hash}.cache"

    def is_rfp_step1_cache_valid(
        self,
        rfp_path: str,
        code_ban_file: Optional[str],
        verbose: bool = False,
        extra_key: Optional[str] = None
    ) -> bool:
        """Проверяет валидность кэша step1: rfp и code_ban_file не менялись."""
        cache_path = self._get_rfp_step1_cache_path(rfp_path, code_ban_file, extra_key)
        if not cache_path.exists():
            return False
        try:
            with open(cache_path, 'rb') as f:
                data = pickle.load(f)
            if not isinstance(data, dict):
                return False
            cache_rfp_mtime = data.get('rfp_mtime')
            cache_code_ban_mtime = data.get('code_ban_mtime', 0)
            if cache_rfp_mtime is None:
                return False
            if os.path.exists(rfp_path) and abs(os.path.getmtime(rfp_path) - cache_rfp_mtime) >= 1e-6:
                return False
            if code_ban_file and os.path.exists(code_ban_file):
                if abs(os.path.getmtime(code_ban_file) - cache_code_ban_mtime) >= 1e-6:
                    return False
            elif code_ban_file and not os.path.exists(code_ban_file):
                return False
            elif not code_ban_file and cache_code_ban_mtime != 0:
                return False
            return True
        except Exception:
            return False

    def _validate_rfp_step1_mtime(
        self,
        data: dict,
        rfp_path: str,
        code_ban_file: Optional[str]
    ) -> bool:
        """Проверяет mtime по уже загруженным данным (без повторного чтения файла)."""
        if not isinstance(data, dict):
            return False
        cache_rfp_mtime = data.get('rfp_mtime')
        cache_code_ban_mtime = data.get('code_ban_mtime', 0)
        if cache_rfp_mtime is None:
            return False
        if os.path.exists(rfp_path) and abs(os.path.getmtime(rfp_path) - cache_rfp_mtime) >= 1e-6:
            return False
        if code_ban_file and os.path.exists(code_ban_file):
            if abs(os.path.getmtime(code_ban_file) - cache_code_ban_mtime) >= 1e-6:
                return False
        elif code_ban_file and not os.path.exists(code_ban_file):
            return False
        elif not code_ban_file and cache_code_ban_mtime != 0:
            return False
        return True

    def load_rfp_step1_from_cache(
        self,
        rfp_path: str,
        code_ban_file: Optional[str],
        verbose: bool = False,
        extra_key: Optional[str] = None
    ) -> Optional[Tuple[List[RowStd], List[RowStd]]]:
        """Загружает полный результат step1 из кэша. Возвращает (rfp_data, error_rows) или None.
        Загружает файл один раз, проверяет mtime по загруженным данным (без повторной загрузки)."""
        cache_path = self._get_rfp_step1_cache_path(rfp_path, code_ban_file, extra_key)
        if not cache_path.exists():
            return None
        try:
            with open(cache_path, 'rb') as f:
                data = pickle.load(f)
            if not self._validate_rfp_step1_mtime(data, rfp_path, code_ban_file):
                return None
            rfp_data = data.get('rfp_data')
            error_rows = data.get('error_rows', [])
            if rfp_data is None:
                return None
            if verbose:
                print(f"    step1 результат загружен из кэша: {cache_path}")
            return (rfp_data, error_rows)
        except Exception as e:
            print(f"Ошибка при загрузке step1 кэша: {e}")
            return None

    def save_rfp_step1_to_cache(
        self,
        rfp_path: str,
        code_ban_file: Optional[str],
        rfp_data: List[RowStd],
        error_rows: List[RowStd],
        verbose: bool = False,
        extra_key: Optional[str] = None
    ) -> bool:
        """Сохраняет полный результат step1 в кэш."""
        try:
            cache_path = self._get_rfp_step1_cache_path(rfp_path, code_ban_file, extra_key)
            rfp_mtime = os.path.getmtime(rfp_path) if os.path.exists(rfp_path) else 0.0
            code_ban_mtime = os.path.getmtime(code_ban_file) if code_ban_file and os.path.exists(code_ban_file) else 0.0
            data = {
                'rfp_mtime': rfp_mtime,
                'code_ban_mtime': code_ban_mtime,
                'rfp_data': rfp_data,
                'error_rows': error_rows,
            }
            with open(cache_path, 'wb') as f:
                pickle.dump(data, f)
            if verbose:
                print(f"    step1 результат сохранён в кэш: {cache_path}")
            return True
        except Exception as e:
            print(f"Ошибка при сохранении step1 кэша: {e}")
            return False

    def get_cache_info(self) -> dict:
        """
        Получает информацию о кэше
        
        Returns:
            Словарь с информацией о кэше
        """
        cache_files = list(self.cache_dir.glob("*.cache"))
        total_size = sum(f.stat().st_size for f in cache_files)
        
        return {
            "cache_dir": str(self.cache_dir),
            "files_count": len(cache_files),
            "total_size_bytes": total_size,
            "total_size_mb": round(total_size / (1024 * 1024), 2),
            "files": [str(f) for f in cache_files]
        }


# Глобальный экземпляр менеджера кэша
cache_manager = CacheManager()


def set_cache_dir(cache_dir: str) -> None:
    """
    Устанавливает директорию кэша для глобального менеджера.
    
    Args:
        cache_dir: Директория для хранения кэш-файлов
    """
    cache_manager.cache_dir = Path(cache_dir)
    cache_manager.cache_dir.mkdir(exist_ok=True)
