# RemitSaver – Outlook Classic COM Add-in

RemitSaver is a pure-Python COM add-in for **Outlook Classic (desktop)** that
automatically extracts remittance amounts from incoming emails and saves the
attachments renamed to `AMOUNT COMPANY.ext` in a folder you choose.

It adds a **RemitSaver** tab to the Outlook ribbon with three buttons:

| Button | Purpose |
|---|---|
| **Extract Now** | Run extraction for selected payers |
| **Manage Payers** | Add/Edit/Delete/Test payer configurations |
| **Settings** | Global settings (save root, log path, startup options) |

---

## Requirements

- Windows 10 / 11
- Outlook Classic (desktop) – **not** Outlook New or Outlook on the web
- Python 3.9+ (64-bit recommended to match Office bitness)
- Run `install.bat` **as Administrator**

---

## Installation

### 1 – Install Python dependencies

```
pip install -r requirements.txt
```

> After `pywin32` installs for the first time, run the post-install script once:
> ```
> python Scripts\pywin32_postinstall.py -install
> ```
> (The `install.bat` does this automatically.)

### 2 – Register the COM add-in

Right-click **`install.bat`** → **Run as administrator**

The script will:
- Verify/install Python dependencies
- Register the COM server via `win32com.server.register`
- Write the Outlook add-in discovery key to:
  `HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect`
- Create `%APPDATA%\RemitSaver\` for config and log files

### 3 – Restart Outlook

Close and re-open Outlook. The **RemitSaver** tab appears in the Explorer
ribbon (the main Outlook window).

> **Tip:** If the tab does not appear, go to
> *File → Options → Add-ins → COM Add-ins → Go…*
> and tick **RemitSaver**.

---

## First-time Setup

1. Click **Manage Payers** → **Add**
2. Fill in:
   - **Name** – e.g. `Nationwide`
   - **Sender filter(s)** – e.g. `@nationwide.com` (comma-separated)
   - **Watch folder** – click *Browse…* to pick an Outlook folder
   - **Save path** – click *Browse…* to pick a local folder
   - **File types** – `all` or `pdf, xlsx`
   - **Amount location** / **Amount strategy** as needed
3. Click **Save**
4. Click **Extract Now**, tick the payer, click **Run Extraction**

---

## File Structure

```
remitsaver/
├── remitsaver_addin.py   # COM server – IDTExtensibility2 + IRibbonExtensibility
├── config.py             # Load/save payers.json and settings.json
├── extractor.py          # Amount/company extraction, Outlook scanning
├── dialogs.py            # All tkinter UI dialogs
├── ribbon.xml            # Ribbon XML (reference copy; content embedded in addin)
├── install.bat           # Register COM add-in (run as Administrator)
├── uninstall.bat         # Unregister COM add-in (run as Administrator)
├── requirements.txt
└── README.md
```

Configuration files written at runtime:

```
%APPDATA%\RemitSaver\
├── payers.json           # Payer configurations
├── settings.json         # Global settings
└── remitsaver.log        # Extraction log
```

---

## How Amount Extraction Works

For each matching email, RemitSaver searches for dollar amounts using the
regex:

```
\$?\s?(\d{1,3}(?:,\d{3})*\.\d{2})|\$?\s?(\d+\.\d{2})
```

Sources searched (in order for **Auto** mode):

1. **Attachments** – PDF (pdfplumber), Excel (openpyxl / xlrd), CSV / TXT
2. **Subject line**
3. **Email body** (first 3,000 characters)

**Amount strategy:**
- `largest` – picks `max()` of all found amounts
- `last` – picks the last match found

**Company name** is derived from the sender's display name by stripping
common corporate suffixes (*Inc, LLC, Corp, LLP, Ltd, Co, N.A., Group,
Holdings*). If the display name is an email address the domain name
(minus TLD) is used and capitalised.

**Filename collision** is handled by appending ` (1)`, ` (2)` etc.

---

## Logging

Every saved file is appended to the log in the format:

```
YYYY-MM-DD HH:MM:SS | INFO | <payer> | <filename> | <sender_email> | <subject>
```

Default log path: `%APPDATA%\RemitSaver\remitsaver.log`

---

## Uninstallation

Right-click **`uninstall.bat`** → **Run as administrator**

Config files in `%APPDATA%\RemitSaver\` are kept. Delete that folder
manually for a full clean.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Tab not visible after restart | File → Options → Add-ins → COM Add-ins → Go → tick RemitSaver |
| Add-in marked as disabled | File → Options → Add-ins → Disabled Items → Enable RemitSaver |
| `import win32com` fails | Run `pip install pywin32` then `python Scripts\pywin32_postinstall.py -install` |
| `pythoncom.CoInitialize` error | Ensure you are not calling COM from multiple threads without `CoInitializeEx` |
| 64-bit Outlook + 32-bit Python | Install 64-bit Python to match Office bitness |

---

## Technical Notes

- **No VSTO, no .NET** – pure Python COM via `pywin32` + `comtypes`
- Implements `IDTExtensibility2` (Outlook loads/unloads) and
  `IRibbonExtensibility` (Ribbon XML delivery)
- UI dialogs run on **daemon threads** so Outlook never freezes
- All file paths handle spaces and Unicode via Python's native `os.path`
- `LoadBehavior = 3` → add-in loads automatically on every Outlook start
