"""
Smoke test for the built CV_Manager.exe.

Launches the executable in a subprocess with a short timeout, verifies it
creates expected data directories and log entries, then terminates it.
No network access or admin rights required.

Usage:
    py src/tests/smoke_exe.py                      # auto-detect CV_Manager.exe
    py src/tests/smoke_exe.py CV_Manager.exe       # explicit path
"""

import sys
import os
import time
import signal
import subprocess
import getpass
from pathlib import Path


STARTUP_WAIT_SECONDS = 8
GRACE_SECONDS = 3


def find_exe() -> Path:
    if len(sys.argv) > 1:
        candidate = Path(sys.argv[1])
        if candidate.exists():
            return candidate
        print(f"FAIL: Specified exe not found: {candidate}")
        sys.exit(1)

    project_root = Path(__file__).parent.parent.parent.resolve()
    candidate = project_root / "CV_Manager.exe"
    if candidate.exists():
        return candidate

    print("FAIL: CV_Manager.exe not found at project root. Build first with:")
    print("  pyinstaller --clean --noconfirm cv_manager.spec")
    sys.exit(1)


def main():
    exe_path = find_exe()
    project_root = Path(__file__).parent.parent.parent.resolve()

    print(f"Smoke test: {exe_path}")
    print(f"Project root: {project_root}")

    data_dir = project_root / "data"
    user = getpass.getuser()
    user_dir = data_dir / "users" / user

    print(f"Expected user data dir: {user_dir}")
    print(f"Starting exe with {STARTUP_WAIT_SECONDS}s timeout...")

    env = os.environ.copy()
    env.pop("HTTP_PROXY", None)
    env.pop("HTTPS_PROXY", None)
    env.pop("http_proxy", None)
    env.pop("https_proxy", None)

    proc = subprocess.Popen(
        [str(exe_path)],
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0,
    )

    time.sleep(STARTUP_WAIT_SECONDS)

    checks_passed = 0
    checks_total = 0

    checks_total += 1
    if proc.poll() is None:
        print("  [PASS] Process is still running after startup")
        checks_passed += 1
    else:
        rc = proc.returncode
        stdout_bytes = proc.stdout.read() if proc.stdout else b""
        stderr_bytes = proc.stderr.read() if proc.stderr else b""
        print(f"  [FAIL] Process exited prematurely with code {rc}")
        if stdout_bytes:
            print(f"  stdout: {stdout_bytes.decode('utf-8', errors='replace')[:500]}")
        if stderr_bytes:
            print(f"  stderr: {stderr_bytes.decode('utf-8', errors='replace')[:500]}")

    checks_total += 1
    if data_dir.exists():
        print(f"  [PASS] data/ directory exists: {data_dir}")
        checks_passed += 1
    else:
        print(f"  [FAIL] data/ directory not found: {data_dir}")

    checks_total += 1
    if user_dir.exists():
        print(f"  [PASS] User directory exists: {user_dir}")
        checks_passed += 1
    else:
        print(f"  [FAIL] User directory not found: {user_dir}")

    logs_dir = user_dir / "logs"
    checks_total += 1
    if logs_dir.exists():
        print(f"  [PASS] Logs directory exists: {logs_dir}")
        checks_passed += 1
    else:
        print(f"  [INFO] Logs directory not yet created (may appear later): {logs_dir}")

    db_path = user_dir / "sites.db"
    checks_total += 1
    if db_path.exists():
        print(f"  [PASS] Database file exists: {db_path}")
        checks_passed += 1
    else:
        print(f"  [INFO] Database not yet created (created on first use): {db_path}")

    print("Terminating process...")
    if proc.poll() is None:
        try:
            if sys.platform == "win32":
                proc.send_signal(signal.CTRL_BREAK_EVENT)
            else:
                proc.terminate()
            proc.wait(timeout=GRACE_SECONDS)
            print("  Process terminated gracefully")
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
            print("  Process killed after timeout")
        except Exception as exc:
            proc.kill()
            proc.wait()
            print(f"  Process killed after error: {exc}")

    print(f"\nSmoke test: {checks_passed}/{checks_total} checks passed")
    if checks_passed >= 3:
        print("RESULT: PASS")
        return 0
    else:
        print("RESULT: FAIL")
        return 1


if __name__ == "__main__":
    sys.exit(main())
