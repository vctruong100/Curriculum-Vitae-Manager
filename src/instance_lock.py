"""
Single-instance lock to prevent multiple GUI instances writing to the same DB.

Uses a lock file in the user data directory. On Windows, msvcrt.locking
provides mandatory file locking. On POSIX, fcntl.flock provides advisory
locking. Both prevent a second process from acquiring the lock while the
first holds it.

Includes automatic stale lock cleanup: if lock acquisition fails, checks
whether the PID in the lock file is still running. If not, removes the
stale lock and retries once.
"""

import sys
import os
import logging

_lock_logger = logging.getLogger(__name__)

_lock_fh = None
_lock_path = None


def _is_process_running(pid: int) -> bool:
    """Check if a process with the given PID is currently running."""
    if sys.platform == "win32":
        import ctypes
        kernel32 = ctypes.windll.kernel32
        PROCESS_QUERY_INFORMATION = 0x0400
        handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False


def _try_acquire_lock(lock_path: str) -> bool:
    """Attempt to acquire the lock file. Returns True on success."""
    global _lock_fh
    try:
        _lock_fh = open(lock_path, "w")
        if sys.platform == "win32":
            import msvcrt
            msvcrt.locking(_lock_fh.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(_lock_fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        _lock_fh.write(str(os.getpid()))
        _lock_fh.flush()
        return True
    except (OSError, IOError):
        if _lock_fh is not None:
            try:
                _lock_fh.close()
            except Exception:
                pass
            _lock_fh = None
        return False


def _read_lock_pid(lock_path: str) -> int:
    """Read the PID from the lock file. Returns -1 if unreadable."""
    try:
        with open(lock_path, "r") as f:
            content = f.read().strip()
            return int(content)
    except Exception:
        return -1


def acquire_instance_lock(lock_dir: str) -> bool:
    global _lock_fh, _lock_path
    _lock_path = os.path.join(lock_dir, ".cv_manager.lock")
    os.makedirs(lock_dir, exist_ok=True)

    if _try_acquire_lock(_lock_path):
        _lock_logger.info("[Lock] Acquired single-instance lock: %s", _lock_path)
        return True

    _lock_logger.warning("[Lock] Failed to acquire lock on first attempt")
    
    if os.path.exists(_lock_path):
        old_pid = _read_lock_pid(_lock_path)
        if old_pid > 0:
            if _is_process_running(old_pid):
                _lock_logger.warning("[Lock] Lock held by running process PID=%d", old_pid)
                return False
            else:
                _lock_logger.info("[Lock] Stale lock detected (PID=%d not running) - removing", old_pid)
                try:
                    os.remove(_lock_path)
                except Exception as exc:
                    _lock_logger.warning("[Lock] Could not remove stale lock: %s", exc)
                    return False
                
                if _try_acquire_lock(_lock_path):
                    _lock_logger.info("[Lock] Acquired lock after stale cleanup: %s", _lock_path)
                    return True
                else:
                    _lock_logger.warning("[Lock] Failed to acquire lock even after cleanup")
                    return False
        else:
            _lock_logger.warning("[Lock] Lock file exists but PID unreadable - manual cleanup needed")
            return False
    
    _lock_logger.warning("[Lock] Lock acquisition failed for unknown reason")
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
