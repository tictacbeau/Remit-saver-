# remitsaver_addin.py
# RemitSaver Outlook COM Add-in
# Implements IDTExtensibility2 + IRibbonExtensibility via pywin32

import sys
import os
import threading
import logging

# ── Ensure the add-in's own directory is on sys.path so sibling imports work ──
_ADDIN_DIR = os.path.dirname(os.path.abspath(__file__))
if _ADDIN_DIR not in sys.path:
    sys.path.insert(0, _ADDIN_DIR)

import win32com.server.register
import pythoncom
import winreg

# ── COM GUIDs ─────────────────────────────────────────────────────────────────
# These are FIXED – do not change after first registration.
ADDIN_PROGID  = "RemitSaver.Connect"
ADDIN_CLASSID = "{6D8B3E2A-4F1C-4A7D-9B5E-2C8F1A3D6E9B}"

# ── Registry key that Outlook reads ──────────────────────────────────────────
OUTLOOK_ADDIN_KEY = r"Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect"

# ── Interface IIDs ────────────────────────────────────────────────────────────
# pywin32's BasicWrapPolicy uses _com_interfaces_ to answer QueryInterface calls.
# Without these, Outlook's QI for IDTExtensibility2 returns E_NOINTERFACE and
# the add-in is never loaded.
IID_IDTExtensibility2   = "{B65AD801-ABAF-11D0-BB8B-0060081897C8}"
IID_IRibbonExtensibility = "{000C0396-0000-0000-C000-000000000046}"

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
      IDTExtensibility2   – Outlook loads/unloads this class as an add-in
      IRibbonExtensibility – supplies the Ribbon XML to Office
    """

    # ── COM registration metadata ──────────────────────────────────────────
    _reg_clsid_  = ADDIN_CLASSID
    _reg_progid_ = ADDIN_PROGID
    _reg_desc_   = "RemitSaver Outlook Add-in"
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER

    # BUG FIX #2: list the interfaces this server supports so that Outlook's
    # QueryInterface calls for IDTExtensibility2 and IRibbonExtensibility
    # succeed instead of returning E_NOINTERFACE.
    _com_interfaces_ = [
        IID_IDTExtensibility2,
        IID_IRibbonExtensibility,
    ]

    # BUG FIX #1: every method Outlook (or the Ribbon host) calls through
    # IDispatch must be listed here.  An empty list means no methods are
    # visible to COM callers, so the add-in loads but nothing ever happens.
    _public_methods_ = [
        # IDTExtensibility2
        "OnConnection",
        "OnDisconnection",
        "OnAddInsUpdate",
        "OnStartupComplete",
        "OnBeginShutdown",
        # IRibbonExtensibility
        "GetCustomUI",
        # Ribbon onLoad / onAction callbacks
        "ribbon_on_load",
        "on_extract_now",
        "on_manage_payers",
        "on_settings",
    ]

    # BUG FIX #3: state must be instance-level, not class-level.
    # Class-level attributes are shared across all instances and are mutated
    # in place, which causes the first instance's state to bleed into any
    # subsequent instance created during the same process lifetime.
    def __init__(self):
        self._application   = None   # Outlook.Application COM object
        self._ribbon_ui     = None   # IRibbonUI handle returned by onLoad
        self._auto_run_done = False  # guard against duplicate startup runs

    # ══ IDTExtensibility2 ════════════════════════════════════════════════════

    def OnConnection(self, application, connect_mode, addin_inst, custom):
        """Called by Outlook when the add-in is loaded."""
        try:
            logger.info("OnConnection called (mode=%s)", connect_mode)
            self._application = application
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
        """Return Ribbon XML; called by Office before the Explorer window opens."""
        logger.debug("GetCustomUI called for ribbon_id=%s", ribbon_id)
        return RIBBON_XML

    # ══ Ribbon callbacks (onLoad / onAction) ═════════════════════════════════

    def ribbon_on_load(self, ribbon_ui):
        """Office fires this once after parsing the Ribbon XML."""
        self._ribbon_ui = ribbon_ui
        logger.info("Ribbon loaded")

    def on_extract_now(self, control):
        logger.info("Extract Now clicked")
        _run_async(self._open_extract_dialog)

    def on_manage_payers(self, control):
        logger.info("Manage Payers clicked")
        _run_async(self._open_manage_payers_dialog)

    def on_settings(self, control):
        logger.info("Settings clicked")
        _run_async(self._open_settings_dialog)

    # ══ Dialog launchers ═════════════════════════════════════════════════════
    # BUG FIX #4: each background thread must call CoInitialize before any
    # COM calls.  Outlook's Application object lives in an STA apartment;
    # accessing it from a thread that has not initialised COM raises
    # "CoInitialize has not been called" (pythoncom.com_error -2147221008).
    #
    # BUG FIX #5: the hidden tk.Tk() root window must be destroyed after the
    # dialog closes, otherwise a ghost window accumulates every time a button
    # is clicked and eventually exhausts GDI/USER handles.

    def _open_extract_dialog(self):
        pythoncom.CoInitialize()
        try:
            import tkinter as tk
            from dialogs import ExtractNowDialog
            root = tk.Tk()
            root.withdraw()
            try:
                dlg = ExtractNowDialog(root, self._application)
                dlg.run()
            finally:
                root.destroy()
        except Exception:
            logger.exception("ExtractNow dialog error")
        finally:
            pythoncom.CoUninitialize()

    def _open_manage_payers_dialog(self):
        pythoncom.CoInitialize()
        try:
            import tkinter as tk
            from dialogs import ManagePayersDialog
            root = tk.Tk()
            root.withdraw()
            try:
                dlg = ManagePayersDialog(root, self._application)
                dlg.run()
            finally:
                root.destroy()
        except Exception:
            logger.exception("ManagePayers dialog error")
        finally:
            pythoncom.CoUninitialize()

    def _open_settings_dialog(self):
        pythoncom.CoInitialize()
        try:
            import tkinter as tk
            from dialogs import SettingsDialog
            root = tk.Tk()
            root.withdraw()
            try:
                dlg = SettingsDialog(root)
                dlg.run()
            finally:
                root.destroy()
        except Exception:
            logger.exception("Settings dialog error")
        finally:
            pythoncom.CoUninitialize()

    def _do_auto_extract(self):
        """Auto-extraction on startup (runs on a background thread)."""
        pythoncom.CoInitialize()
        try:
            from config import load_payers, load_settings
            from extractor import run_extraction_for_payer
            payers   = load_payers()
            settings = load_settings()
            for payer in payers:
                run_extraction_for_payer(payer, self._application, settings)
        except Exception:
            logger.exception("Auto-extract error")
        finally:
            pythoncom.CoUninitialize()


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
        winreg.SetValueEx(key, "Description",    0, winreg.REG_SZ,    "RemitSaver – Remittance extraction add-in")
        winreg.SetValueEx(key, "FriendlyName",   0, winreg.REG_SZ,    "RemitSaver")
        winreg.SetValueEx(key, "LoadBehavior",   0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(key, "CommandLineSafe",0, winreg.REG_DWORD, 0)
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
    _write_outlook_addin_registry()


def DllUnregisterServer():
    _remove_outlook_addin_registry()


# ══ Entry-point ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import win32com.server.register as _reg

    if "--unregister" in sys.argv or "--unregserver" in sys.argv:
        _reg.UnregisterClasses(RemitSaverAddin, quiet=True)
        _remove_outlook_addin_registry()
        print("RemitSaver unregistered.")
    else:
        _reg.RegisterClasses(RemitSaverAddin, quiet=True)
        _write_outlook_addin_registry()
        print("RemitSaver registered successfully.")
        print(f"  ProgID  : {ADDIN_PROGID}")
        print(f"  CLSID   : {ADDIN_CLASSID}")
        print(f"  Log     : {_log_path}")
