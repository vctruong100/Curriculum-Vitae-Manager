import time
import logging
from typing import List, Optional, Dict, Any

from config import UNDO_TIMEOUT_SECONDS


logger = logging.getLogger(__name__)


class UndoBuffer:

    def __init__(self, timeout_seconds: int = UNDO_TIMEOUT_SECONDS):
        self._timeout = timeout_seconds
        self._site_id: Optional[int] = None
        self._studies: List[Dict[str, Any]] = []
        self._timestamp: float = 0.0

    @property
    def can_undo(self) -> bool:
        if not self._studies:
            return False
        if time.monotonic() - self._timestamp > self._timeout:
            logger.info("[UndoBuffer] Buffer expired (%.0fs elapsed)", time.monotonic() - self._timestamp)
            self.clear()
            return False
        return True

    @property
    def site_id(self) -> Optional[int]:
        return self._site_id

    @property
    def study_count(self) -> int:
        return len(self._studies)

    def store(self, site_id: int, studies: List[Dict[str, Any]]) -> None:
        if not studies:
            return
        self._site_id = site_id
        self._studies = list(studies)
        self._timestamp = time.monotonic()
        logger.info(
            "[UndoBuffer] Stored %d study(ies) for site_id=%d",
            len(studies), site_id,
        )

    def pop(self) -> List[Dict[str, Any]]:
        if not self.can_undo:
            return []
        studies = self._studies
        logger.info(
            "[UndoBuffer] Popping %d study(ies) for site_id=%d",
            len(studies), self._site_id,
        )
        self._studies = []
        self._timestamp = 0.0
        return studies

    def clear(self) -> None:
        if self._studies:
            logger.info("[UndoBuffer] Clearing buffer (had %d entries)", len(self._studies))
        self._studies = []
        self._timestamp = 0.0
        self._site_id = None

    def clear_if_site_changed(self, new_site_id: Optional[int]) -> None:
        if self._site_id is not None and self._site_id != new_site_id:
            logger.info(
                "[UndoBuffer] Site changed from %d to %s — clearing",
                self._site_id, new_site_id,
            )
            self.clear()
