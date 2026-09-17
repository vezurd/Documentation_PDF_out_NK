"""Qt thread that pumps a child-process MTO compare without holding the GIL."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.json_line_thread import JsonLineWorkerThread
from rd_catalog.mto_compare_job import config_to_mto_job_dict


class MtoCompareThread(JsonLineWorkerThread):
    """Start ``python -m rd_catalog.mto_compare_cli`` and forward JSON lines."""

    _cli_module = "rd_catalog.mto_compare_cli"
    _job_prefix = "mto_job"
    _cancel_prefix = "mto_cancel"
    _cancel_log = "Запрошена отмена сверки MTO…"
    _process_failure = "Процесс сверки MTO завершился"

    def __init__(
        self,
        config: CatalogConfig,
        *,
        document_keys: Iterable[tuple[str, str]] | None = None,
        priority_keys: Iterable[tuple[str, str]] = (),
        batch_size: int = 1,
        parent=None,
    ) -> None:
        """Store an immutable MTO compare request.

        Args:
            config: Resolved catalog configuration.
            document_keys: Optional MTO keys to restrict work; ``None`` means
                full incremental compare (serialized as JSON null).
            priority_keys: Keys from the latest scan walk; sorted first in plan.
            batch_size: Number of plan pairs per worker batch.
            parent: Optional Qt parent.
        """

        super().__init__(config, batch_size=batch_size, parent=parent)
        self._document_keys = (
            None
            if document_keys is None
            else tuple((str(a), str(b)) for a, b in document_keys)
        )
        self._priority_keys = tuple((str(a), str(b)) for a, b in priority_keys)
        self.remaining_keys: list[tuple[str, str]] = []

    def _job_payload(self, cancel_path: str) -> dict[str, Any]:
        return config_to_mto_job_dict(
            self._config,
            document_keys=(
                None if self._document_keys is None else list(self._document_keys)
            ),
            priority_keys=self._priority_keys,
            cancel_path=cancel_path,
            batch_size=self._batch_size,
        )

    def _handle_done(self, payload: dict[str, Any]) -> None:
        raw_keys = payload.get("remaining_keys") or []
        self.remaining_keys = [
            (str(key[0]), str(key[1]))
            for key in raw_keys
            if isinstance(key, (list, tuple)) and len(key) == 2
        ]
