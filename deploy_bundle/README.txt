BOM Tools - Offline Deployment for Windows
============================================

This package contains a one-click deployment of the BOM Tools web application,
designed for Windows servers WITHOUT internet access.

PREREQUISITES
  - Windows 10/11 or Windows Server 2016+
  - Python 3.10+ installed
    - Download from: https://www.python.org/downloads/
    - IMPORTANT: During installation, CHECK "Add Python to PATH"
      (it's at the bottom of the first installer screen)
    - Do NOT use the Microsoft Store version of Python

HOW TO USE
  1. Copy the entire `deploy_bundle` folder to the target Windows machine
  2. Double-click `install_and_run.bat`
  3. The script will:
     - Automatically detect Python (supports both `python` and `py` commands)
     - Create a virtual environment
     - Install all dependencies from local files (no internet needed)
     - Start the web server at http://localhost:5000
  4. Open http://localhost:5000 in a browser on the server
     Other computers in the LAN can open:
       http://SERVER_IP:5000

SERVER CONFIGURATION
  - Default bind address: 0.0.0.0
  - Default port: 5000
  - To use another port, set PORT before starting:
      set PORT=8080
      install_and_run.bat
  - If other computers need to access this service, allow the selected port
    through Windows Defender Firewall.
  - This package installs Waitress and starts the app with a production WSGI
    server by default. It is suitable for multiple users on the LAN.

WHAT'S INCLUDED
  - Flask web application (web_app2/)
  - All Python dependencies as .whl files (wheels/)
    - Supports Python 3.10, 3.11, 3.12, 3.14
    - Includes Flask, openpyxl, requests and waitress
  - Automatic virtual environment setup
  - Zero internet required
  - Bug reports are stored under web_app2/bug_reports/ on the deployed server.
    Keep that folder if you move or back up the application later.

TROUBLESHOOTING
  - "'python' not recognized" / "Python not found"
    -> Reinstall Python from python.org, check "Add Python to PATH"
  - "Failed to create virtual environment"
    -> You may have the Microsoft Store Python. Uninstall it and
       install Python from https://www.python.org/downloads/
  - Server fails to start
    -> Check if port 5000 is already in use by another application
  - For firewall issues, allow Python through Windows Defender Firewall
  - Other computers cannot open the site
    -> Use the server's LAN IP instead of localhost, for example:
       http://192.168.1.100:5000
    -> Check Windows Defender Firewall inbound rules for the selected port
