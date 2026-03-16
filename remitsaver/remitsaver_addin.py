# remitsaver_addin.py
# RemitSaver Outlook COM Add-in
# Implements IDTExtensibility2 + IRibbonExtensibility via pywin32/comtypes

import sys
import os
import threading
import traceback
import logging

# ── Ensure the add-in's own directory is on sys.path so sibling imports work ──
_ADDIN_DIR = os.path.dirname(os.path.abspath(__file__))
if _ADDIN_DIR not in sys.path:
    sys.path.insert(0, _ADDIN_DIR)

import win32com.client
import win32com.server.register
import pythoncom
import pywintypes
import winreg

# ── COM GUIDs ─────────────────────────────────────────────────────────────────
# These are FIXED – do not change after first registration.
ADDIN_PROGID  = "RemitSaver.Connect"
ADDIN_CLASSID = "{6D8B3E2A-4F1C-4A7D-9B5E-2C8F1A3D6E9B}"

# ── Registry key that Outlook reads ──────────────────────────────────────────
OUTLOOK_ADDIN_KEY = r"Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect"

# ── Ribbon XML ────────────────────────────────────────────────────────────────
RIBBON_XML = r"""<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="ribbon_on_load">
  <ribbon>
    <tabs>
      <tab id="tabRemitSaver" label="RemitSaver">
        <group id="grpMain" label="Remittance">
          <button id="btnExtractNow"
                  label="Extract Now"
                  imageMso="ImportExportDataAccess"
                  size="large"
                  onAction="on_extract_now"
                  screentip="Run remittance extraction for selected payers"
                  supertip="Scans configured Outlook folders and saves attachments renamed to AMOUNT COMPANY.ext"/>
          <button id="btnManagePayers"
                  label="Manage Payers"
                  imageMso="TableInsert"
                  size="large"
                  onAction="on_manage_payers"
                  screentip="Add, edit or remove payer configurations"
                  supertip="Open the payer management dialog to configure sender filters, folders, and extraction settings"/>
          <button id="btnSettings"
                  label="Settings"
                  imageMso="OptionsCustomizeKeyboard"
                  size="large"
                  onAction="on_settings"
                  screentip="Open global RemitSaver settings"
                  supertip="Configure default save path, log location, and startup options"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

# ── Logging bootstrap ─────────────────────────────────────────────────────────
def _get_log_path():
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    log_dir = os.path.join(appdata, "RemitSaver")
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, "remitsaver.log")

_log_path = _get_log_path()
logging.basicConfig(
    filename=_log_path,
    level=logging.DEBUG,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("RemitSaver")


# ── Helper: run a callable on a daemon thread (keeps Outlook responsive) ──────
def _run_async(fn, *args, **kwargs):
    t = threading.Thread(target=fn, args=args, kwargs=kwargs, daemon=True)
    t.start()


# ── The COM server class ───────────────────────────────────────────────────────
class RemitSaverAddin:
    """
    Implements:
      IDTExtensibility2  – Outlook loads/unloads this class as an add-in
      IRibbonExtensibility – supplies the Ribbon XML to Office
    """

    # ── COM registration metadata ──────────────────────────────────────────
    _reg_clsid_     = ADDIN_CLASSID
    _reg_progid_    = ADDIN_PROGID
    _reg_desc_      = "RemitSaver Outlook Add-in"
    _reg_clsctx_    = pythoncom.CLSCTX_INPROC_SERVER
    _reg_policy_    = None          # use default BasicWrapPolicy
    _public_methods_ = []
    _readonly_attrs_ = []

    # IDTExtensibility2 interface GUID
    _com_interfaces_ = []           # pywin32 fills from IIDs below
    _typelib_guid_   = None

    # ── IDTExtensibility2 interface IIDs that Outlook will QI for ─────────
    # We implement them as plain methods; pywin32's BasicWrapPolicy exposes them.
    _iid_IDTExtensibility2_ = "{B65AD801-ABAF-11D0-BB8B-0060081897C8}"

    # ── State ──────────────────────────────────────────────────────────────
    _application    = None   # Outlook.Application
    _ribbon_ui      = None   # IRibbonUI callback handle
    _auto_run_done  = False

    # ══ IDTExtensibility2 ════════════════════════════════════════════════════

    def OnConnection(self, application, connect_mode, addin_inst, custom):
        """Called by Outlook when the add-in is loaded."""
        try:
            logger.info("OnConnection called (mode=%s)", connect_mode)
            self._application = application
            # Trigger Ribbon load; OnStartupComplete/OnBeginShutdown handle the rest
        except Exception:
            logger.exception("OnConnection failed")

    def OnDisconnection(self, mode, custom):
        """Called by Outlook when the add-in is unloaded."""
        logger.info("OnDisconnection (mode=%s)", mode)
        self._application = None
        self._ribbon_ui   = None

    def OnAddInsUpdate(self, custom):
        pass

    def OnStartupComplete(self, custom):
        """Called after Outlook has fully started up."""
        logger.info("OnStartupComplete")
        try:
            from config import load_settings
            settings = load_settings()
            if settings.get("auto_run_on_startup") and not self._auto_run_done:
                self._auto_run_done = True
                _run_async(self._do_auto_extract)
        except Exception:
            logger.exception("OnStartupComplete error")

    def OnBeginShutdown(self, custom):
        logger.info("OnBeginShutdown")

    # ══ IRibbonExtensibility ═════════════════════════════════════════════════

    def GetCustomUI(self, ribbon_id):
        """Return Ribbon XML for the Explorer window ribbon."""
        logger.debug("GetCustomUI called for ribbon_id=%s", ribbon_id)
        return RIBBON_XML

    # ══ Ribbon callbacks (onLoad / onAction) ═════════════════════════════════

    def ribbon_on_load(self, ribbon_ui):
        """Fired once Office has loaded the Ribbon; save the IRibbonUI handle."""
        self._ribbon_ui = ribbon_ui
        logger.info("Ribbon loaded")

    def on_extract_now(self, control):
        """'Extract Now' button clicked."""
        logger.info("Extract Now clicked")
        _run_async(self._open_extract_dialog)

    def on_manage_payers(self, control):
        """'Manage Payers' button clicked."""
        logger.info("Manage Payers clicked")
        _run_async(self._open_manage_payers_dialog)

    def on_settings(self, control):
        """'Settings' button clicked."""
        logger.info("Settings clicked")
        _run_async(self._open_settings_dialog)

    # ══ Dialog launchers (run on background threads) ══════════════════════════

    def _open_extract_dialog(self):
        try:
            import tkinter as tk
            from dialogs import ExtractNowDialog
            root = tk.Tk()
            root.withdraw()
            dlg = ExtractNowDialog(root, self._application)
            dlg.run()
        except Exception:
            logger.exception("ExtractNow dialog error")

    def _open_manage_payers_dialog(self):
        try:
            import tkinter as tk
            from dialogs import ManagePayersDialog
            root = tk.Tk()
            root.withdraw()
            dlg = ManagePayersDialog(root, self._application)
            dlg.run()
        except Exception:
            logger.exception("ManagePayers dialog error")

    def _open_settings_dialog(self):
        try:
            import tkinter as tk
            from dialogs import SettingsDialog
            root = tk.Tk()
            root.withdraw()
            dlg = SettingsDialog(root)
            dlg.run()
        except Exception:
            logger.exception("Settings dialog error")

    def _do_auto_extract(self):
        """Auto-extraction on startup (if enabled in settings)."""
        try:
            from config import load_payers, load_settings
            from extractor import run_extraction_for_payer
            payers   = load_payers()
            settings = load_settings()
            for payer in payers:
                run_extraction_for_payer(payer, self._application, settings)
        except Exception:
            logger.exception("Auto-extract error")


# ══ Registry helpers ══════════════════════════════════════════════════════════

def _write_outlook_addin_registry():
    """Write the HKCU Outlook add-in keys so Outlook discovers RemitSaver."""
    try:
        key = winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            OUTLOOK_ADDIN_KEY,
            0,
            winreg.KEY_WRITE,
        )
        winreg.SetValueEx(key, "Description",   0, winreg.REG_SZ,    "RemitSaver – Remittance extraction add-in")
        winreg.SetValueEx(key, "FriendlyName",  0, winreg.REG_SZ,    "RemitSaver")
        winreg.SetValueEx(key, "LoadBehavior",  0, winreg.REG_DWORD, 3)  # 3 = load at startup
        winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)
        print("[RemitSaver] Outlook add-in registry key written.")
    except Exception as exc:
        print(f"[RemitSaver] Failed to write Outlook registry key: {exc}")


def _remove_outlook_addin_registry():
    """Remove the HKCU Outlook add-in keys."""
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, OUTLOOK_ADDIN_KEY)
        print("[RemitSaver] Outlook add-in registry key removed.")
    except FileNotFoundError:
        pass
    except Exception as exc:
        print(f"[RemitSaver] Failed to remove Outlook registry key: {exc}")


# ══ COM Registration hooks ════════════════════════════════════════════════════

def DllRegisterServer():
    """Called by win32com.server.register when registering."""
    _write_outlook_addin_registry()


def DllUnregisterServer():
    """Called by win32com.server.register when unregistering."""
    _remove_outlook_addin_registry()


# ══ Entry-point ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import win32com.server.register as _reg

    if "--unregister" in sys.argv or "--unregserver" in sys.argv:
        _reg.UnregisterClasses(RemitSaverAddin, quiet=True)
        _remove_outlook_addin_registry()
        print("RemitSaver unregistered.")
    else:
        # Default: register
        _reg.RegisterClasses(RemitSaverAddin, quiet=True)
        _write_outlook_addin_registry()
        print("RemitSaver registered successfully.")
        print(f"  ProgID  : {ADDIN_PROGID}")
        print(f"  CLSID   : {ADDIN_CLASSID}")
        print(f"  Log     : {_log_path}")
