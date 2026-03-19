"""
Create Windows shortcuts (.lnk) with an explicit AppUserModelID.

After a successful build, this script creates:
  - A Desktop shortcut   (CV_Manager.lnk)
  - A Start Menu shortcut (CV_Manager.lnk)

Both shortcuts carry the same AppUserModelID that the running application sets
via SetCurrentProcessExplicitAppUserModelID, so pinned/unpinned shortcuts
group correctly with the running window on the Windows taskbar.

Uses pure ctypes COM — no pywin32 dependency.
No network requests.
"""

import sys
import os
import ctypes
import ctypes.wintypes
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# Must match the value in src/appid.py
APP_USER_MODEL_ID = "CenExel.CVResearchExperienceManager"

# COM CLSIDs / IIDs
CLSID_ShellLink = ctypes.wintypes.GUID("{00021401-0000-0000-C000-000000000046}")
IID_IShellLinkW = ctypes.wintypes.GUID("{000214F9-0000-0000-C000-000000000046}")
IID_IPersistFile = ctypes.wintypes.GUID("{0000010B-0000-0000-C000-000000000046}")
IID_IPropertyStore = ctypes.wintypes.GUID("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")

PKEY_AppUserModel_ID = (
    ctypes.wintypes.GUID("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}"),
    5,
)

# KNOWNFOLDERID for Desktop and Start Menu Programs
FOLDERID_Desktop = ctypes.wintypes.GUID("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}")
FOLDERID_Programs = ctypes.wintypes.GUID("{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}")

VT_LPWSTR = 31


def _get_known_folder(folder_id):
    """Get a known folder path via SHGetKnownFolderPath."""
    buf = ctypes.c_wchar_p()
    hr = ctypes.windll.shell32.SHGetKnownFolderPath(
        ctypes.byref(folder_id),
        0,
        None,
        ctypes.byref(buf),
    )
    if hr != 0:
        return None
    path = buf.value
    ctypes.windll.ole32.CoTaskMemFree(buf)
    return path


def create_shortcut_with_appid(target_exe: str, appid: str, shortcut_path: str,
                                icon_path: str = None, description: str = "") -> bool:
    """Create a .lnk shortcut with an embedded AppUserModelID.

    Parameters
    ----------
    target_exe : str
        Absolute path to the target executable.
    appid : str
        The AppUserModelID to embed in the shortcut.
    shortcut_path : str
        Absolute path where the .lnk file will be written.
    icon_path : str, optional
        Path to the icon file.  Defaults to target_exe.
    description : str, optional
        Shortcut description / tooltip text.

    Returns
    -------
    bool
        True on success.
    """
    if sys.platform != "win32":
        logger.debug("Not on Windows — skipping shortcut creation.")
        return False

    try:
        ctypes.windll.ole32.CoInitialize(None)

        # Create IShellLink instance
        shell_link = ctypes.POINTER(ctypes.c_void_p)()
        hr = ctypes.windll.ole32.CoCreateInstance(
            ctypes.byref(CLSID_ShellLink),
            None,
            1,  # CLSCTX_INPROC_SERVER
            ctypes.byref(IID_IShellLinkW),
            ctypes.byref(shell_link),
        )
        if hr != 0:
            logger.warning("CoCreateInstance(ShellLink) failed: 0x%08X", hr & 0xFFFFFFFF)
            return False

        # Get vtable pointer
        vtbl = ctypes.cast(
            ctypes.cast(shell_link, ctypes.POINTER(ctypes.c_void_p))[0],
            ctypes.POINTER(ctypes.c_void_p * 21),
        ).contents

        # IShellLinkW::SetPath (index 20)
        set_path = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p)(
            vtbl[20]
        )
        set_path(shell_link, target_exe)

        # IShellLinkW::SetWorkingDirectory (index 9)
        set_workdir = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p)(
            vtbl[9]
        )
        set_workdir(shell_link, str(Path(target_exe).parent))

        # IShellLinkW::SetDescription (index 7)
        if description:
            set_desc = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p)(
                vtbl[7]
            )
            set_desc(shell_link, description)

        # IShellLinkW::SetIconLocation (index 17)
        if icon_path:
            set_icon = ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int
            )(vtbl[17])
            set_icon(shell_link, icon_path, 0)

        # QueryInterface for IPropertyStore to set AppUserModelID
        prop_store = ctypes.c_void_p()
        qi = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.POINTER(ctypes.wintypes.GUID), ctypes.POINTER(ctypes.c_void_p)
        )(vtbl[0])
        hr = qi(shell_link, ctypes.byref(IID_IPropertyStore), ctypes.byref(prop_store))

        if hr == 0 and prop_store:
            ps_vtbl = ctypes.cast(
                ctypes.cast(prop_store, ctypes.POINTER(ctypes.c_void_p))[0],
                ctypes.POINTER(ctypes.c_void_p * 8),
            ).contents

            # Build PROPVARIANT with VT_LPWSTR
            class PROPVARIANT(ctypes.Structure):
                _fields_ = [
                    ("vt", ctypes.c_ushort),
                    ("reserved1", ctypes.c_ushort),
                    ("reserved2", ctypes.c_ushort),
                    ("reserved3", ctypes.c_ushort),
                    ("pwszVal", ctypes.c_wchar_p),
                    ("pad", ctypes.c_void_p),
                ]

            pv = PROPVARIANT()
            pv.vt = VT_LPWSTR
            pv.pwszVal = appid

            # Build PROPERTYKEY
            class PROPERTYKEY(ctypes.Structure):
                _fields_ = [
                    ("fmtid", ctypes.wintypes.GUID),
                    ("pid", ctypes.wintypes.DWORD),
                ]

            pk = PROPERTYKEY()
            pk.fmtid = PKEY_AppUserModel_ID[0]
            pk.pid = PKEY_AppUserModel_ID[1]

            # IPropertyStore::SetValue (index 6)
            set_value = ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p,
                ctypes.POINTER(PROPERTYKEY), ctypes.POINTER(PROPVARIANT),
            )(ps_vtbl[6])
            set_value(prop_store, ctypes.byref(pk), ctypes.byref(pv))

            # IPropertyStore::Commit (index 7)
            commit = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(ps_vtbl[7])
            commit(prop_store)

            # Release IPropertyStore
            release_ps = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(ps_vtbl[2])
            release_ps(prop_store)
            logger.info("Set AppUserModelID=%r on shortcut.", appid)

        # QueryInterface for IPersistFile and save
        persist_file = ctypes.c_void_p()
        hr = qi(shell_link, ctypes.byref(IID_IPersistFile), ctypes.byref(persist_file))
        if hr != 0:
            logger.warning("QueryInterface(IPersistFile) failed: 0x%08X", hr & 0xFFFFFFFF)
            return False

        pf_vtbl = ctypes.cast(
            ctypes.cast(persist_file, ctypes.POINTER(ctypes.c_void_p))[0],
            ctypes.POINTER(ctypes.c_void_p * 8),
        ).contents

        # IPersistFile::Save (index 6)
        save = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int
        )(pf_vtbl[6])
        hr = save(persist_file, shortcut_path, 1)

        # Release
        release_pf = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(pf_vtbl[2])
        release_pf(persist_file)

        release_sl = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtbl[2])
        release_sl(shell_link)

        if hr == 0:
            logger.info("Shortcut created: %s", shortcut_path)
            return True
        else:
            logger.warning("IPersistFile::Save failed: 0x%08X", hr & 0xFFFFFFFF)
            return False

    except Exception as exc:
        logger.warning("Shortcut creation failed: %s", exc)
        return False
    finally:
        try:
            ctypes.windll.ole32.CoUninitialize()
        except Exception:
            pass


def main():
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    project_root = Path(__file__).parent.parent.resolve()

    # Determine the exe path
    onefile_exe = project_root / "CV_Manager.exe"
    onedir_exe = project_root / "CV_Manager" / "CV_Manager.exe"

    if onefile_exe.exists():
        target_exe = str(onefile_exe)
    elif onedir_exe.exists():
        target_exe = str(onedir_exe)
    else:
        logger.warning("No built exe found — skipping shortcut creation.")
        sys.exit(0)

    icon_path = str(project_root / "build" / "assets" / "app.ico")
    if not Path(icon_path).exists():
        icon_path = target_exe

    description = "CV Research Experience Manager"

    # Desktop shortcut
    desktop = _get_known_folder(FOLDERID_Desktop)
    if desktop:
        lnk = os.path.join(desktop, "CV_Manager.lnk")
        ok = create_shortcut_with_appid(target_exe, APP_USER_MODEL_ID, lnk,
                                         icon_path=icon_path, description=description)
        if ok:
            logger.info("Desktop shortcut: %s", lnk)

    # Start Menu shortcut
    programs = _get_known_folder(FOLDERID_Programs)
    if programs:
        lnk = os.path.join(programs, "CV_Manager.lnk")
        ok = create_shortcut_with_appid(target_exe, APP_USER_MODEL_ID, lnk,
                                         icon_path=icon_path, description=description)
        if ok:
            logger.info("Start Menu shortcut: %s", lnk)


if __name__ == "__main__":
    main()
