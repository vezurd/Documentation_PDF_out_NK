"""Qt thread that pumps a child-process export MTO pair compare."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.json_line_thread import JsonLineWorkerThread
from rd_catalog.mto_pair_compare import MtoFilePair, PairComparisonResult
from rd_catalog.mto_pair_compare_job import (
    config_to_pair_job_dict,
    pairs_from_payload,
    session_verdicts_from_payload,
)


class MtoPairCompareThread(JsonLineWorkerThread):
    """Start ``python -m rd_catalog.mto_pair_compare_cli`` and forward JSON lines.

    Shares process pumping, job-file lifetime, BELOW_NORMAL priority, and
    flag-file cancel with :class:`rd_catalog.mto_compare_thread.MtoCompareThread`
    via :class:`rd_catalog.json_line_thread.JsonLineWorkerThread`.
    """

    _cli_module = "rd_catalog.mto_pair_compare_cli"
    _job_prefix = "mto_pair_job"
    _cancel_prefix = "mto_pair_cancel"
    _cancel_log = "Запрошена отмена сверки пар MTO…"
    _process_failure = "Процесс сверки пар MTO завершился"

    def __init__(
        self,
        config: CatalogConfig,
        *,
        pairs: Iterable[MtoFilePair] = (),
        batch_size: int = 1,
        parent=None,
    ) -> None:
        """Store an immutable export-pair compare request.

        Args:
            config: Resolved catalog configuration.
            pairs: Source/destination pairs to consider.
            batch_size: Number of pending pairs per worker batch.
            parent: Optional Qt parent.
        """

        super().__init__(config, batch_size=batch_size, parent=parent)
        self._pairs = tuple(pairs)
        self.remaining_pairs: list[MtoFilePair] = []
        self.session_verdicts: dict[tuple[str, str], PairComparisonResult] = {}

    def _job_payload(self, cancel_path: str) -> dict[str, Any]:
        return config_to_pair_job_dict(
            self._config,
            pairs=self._pairs,
            cancel_path=cancel_path,
            batch_size=self._batch_size,
        )

    def _handle_done(self, payload: dict[str, Any]) -> None:
        self.remaining_pairs = pairs_from_payload(payload.get("remaining_pairs"))
        self.session_verdicts = session_verdicts_from_payload(
            payload.get("session_verdicts")
        )
