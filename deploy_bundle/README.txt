BOM Tools - Offline LAN Deployment for Windows
===============================================

This folder is the offline deployment bundle for the BOM Tools web app.
Copy the whole deploy_bundle folder to the target LAN server and run it there.
No internet connection is required if Python is already installed.

PREREQUISITES
  - Windows 10/11 or Windows Server 2016+
  - Python 3.10, 3.11, 3.12 or 3.14 installed from python.org
  - During Python installation, CHECK "Add Python to PATH"
  - Do NOT use the Microsoft Store Python version

QUICK START
  1. Copy the entire deploy_bundle folder to the target server.
  2. Double-click install_and_run.bat.
  3. The script will create venv, install packages from wheels/, and start:
       http://localhost:5000
  4. Other LAN users open:
       http://SERVER_IP:5000

CHANGE PORT
  In cmd, run:
      set PORT=8080
      install_and_run.bat

FIREWALL
  If other computers cannot open the site, allow inbound TCP traffic for the
  selected port, default 5000, in Windows Defender Firewall.

DATA DIRECTORIES
  Runtime data is created on the deployed server under web_app2/:
  - bug_reports/        Bug records, status changes, and attachments
  - feature_requests/   Demand/work-order records, likes, and attachments
  - manufacturer_aliases/ Manufacturer alias mapping SQLite database
  - uploads/            Temporary uploaded Excel files
  - outputs/            Generated result files
  - cache/              Feishu cache data

Keep bug_reports/, feature_requests/, and manufacturer_aliases/ when backing up or moving the server.
Temporary uploads/outputs/cache/ can be cleaned if needed.

INCLUDED TOOLS
  - BOM format conversion
  - Feishu preferred-library and relation-library matching
  - Manufacturer naming alias map
  - BOM preferred-rate query
  - PLM upload format conversion
  - PLM web automation for spec reverse material lookup
  - BOM compare tool collection
  - Bug submission with status management
  - Demand development work orders with likes

OFFLINE DEPENDENCIES
  wheels/ includes Flask, openpyxl, requests, playwright, waitress, and transitive deps.
  The installer uses:
      pip install --no-index --find-links wheels -r requirements.txt

PLAYWRIGHT CHROMIUM
  PLM web automation needs the Playwright Chromium browser runtime.
  Online target server:
      venv\Scripts\python.exe -m playwright install chromium
  Offline target server:
      Include a prepared ms-playwright Chromium cache with the bundle, or set
      PLAYWRIGHT_BROWSERS_PATH to the folder that contains the Chromium cache.
  Without this browser runtime, the normal BOM tools still start, but PLM web
  automation will fail when launching Chromium.

TROUBLESHOOTING
  - Python not found:
      Reinstall Python from python.org and select "Add Python to PATH".
  - venv creation failed:
      Remove Microsoft Store Python and install python.org Python.
  - Port already in use:
      Set another PORT before running install_and_run.bat.
  - LAN users cannot open:
      Use the server LAN IP, not localhost, and check firewall inbound rules.
  - PLM web automation cannot launch Chromium:
      Install Chromium with Playwright or provide the offline ms-playwright cache.
