"""
Single-instance lock to prevent multiple GUI instances writing to the same DB.

Uses a lock file in the user data directory. On Windows, msvcrt.locking
provides mandatory file locking. On POSIX, fcntl.flock provides advisory
locking. Both prevent a second process from acquiring the lock while the
first holds it.
"""

import sys
import os
import logging

_lock_logger = logging.getLogger(__name__)

_lock_fh = None
_lock_path = None


def acquire_instance_lock(lock_dir: str) -> bool:
    global _lock_fh, _lock_path
    _lock_path = os.path.join(lock_dir, ".cv_manager.lock")
    os.makedirs(lock_dir, exist_ok=True)

    try:
        _lock_fh = open(_lock_path, "w")
        if sys.platform == "win32":
            import msvcrt
            msvcrt.locking(_lock_fh.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(_lock_fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        _lock_fh.write(str(os.getpid()))
        _lock_fh.flush()
        _lock_logger.info("[Lock] Acquired single-instance lock: %s", _lock_path)
        return True
    except (OSError, IOError) as exc:
        _lock_logger.warning("[Lock] Failed to acquire lock: %s", exc)
        if _lock_fh is not None:
            try:
                _lock_fh.close()
            except Exception:
                pass
            _lock_fh = None
        return False


def release_instance_lock() -> None:
    global _lock_fh, _lock_path
    if _lock_fh is None:
        return
    try:
        if sys.platform == "win32":
            import msvcrt
            try:
                _lock_fh.seek(0)
                msvcrt.locking(_lock_fh.fileno(), msvcrt.LK_UNLCK, 1)
            except Exception:
                pass
        else:
            import fcntl
            fcntl.flock(_lock_fh, fcntl.LOCK_UN)
        _lock_fh.close()
        _lock_logger.info("[Lock] Released single-instance lock")
    except Exception as exc:
        _lock_logger.debug("[Lock] Error releasing lock: %s", exc)
    finally:
        _lock_fh = None
    if _lock_path is not None:
        try:
            os.remove(_lock_path)
        except Exception:
            pass
        _lock_path = None
