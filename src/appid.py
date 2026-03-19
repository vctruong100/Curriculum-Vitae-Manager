"""
AppUserModelID management for CV Research Experience Manager.

Windows uses AppUserModelID to group taskbar icons.  When a Python / PyInstaller
application does not set an explicit ID, Windows falls back to the default Python
ID and may show a second taskbar icon for pinned shortcuts.

Call set_app_user_model_id() as early as possible in the process — before any
window is created — so that pinned shortcuts and running windows share the same
identity.

No network requests.  Purely local Win32 API call.
"""

import sys
import logging

logger = logging.getLogger(__name__)

APP_USER_MODEL_ID = "CenExel.CVResearchExperienceManager"


def set_app_user_model_id(app_id: str = None) -> bool:
    """Set the Windows AppUserModelID for this process.

    Parameters
    ----------
    app_id : str, optional
        The AppUserModelID string.  Defaults to APP_USER_MODEL_ID.

    Returns
    -------
    bool
        True if the ID was set successfully, False otherwise.
    """
    if sys.platform != "win32":
        logger.debug("[AppID] Not on Windows — skipping AppUserModelID.")
        return False

    if app_id is None:
        app_id = APP_USER_MODEL_ID

    try:
        import ctypes
        result = ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            app_id
        )
        if result == 0:
            logger.info("[AppID] Set AppUserModelID to %r", app_id)
            return True
        else:
            logger.warning(
                "[AppID] SetCurrentProcessExplicitAppUserModelID returned %d",
                result,
            )
            return False
    except AttributeError:
        logger.debug("[AppID] windll.shell32 not available (non-Windows?).")
        return False
    except Exception as exc:
        logger.warning("[AppID] Failed to set AppUserModelID: %s", exc)
        return False
