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
  4. Open http://localhost:5000 in a browser

WHAT'S INCLUDED
  - Flask web application (web_app2/)
  - All Python dependencies as .whl files (wheels/)
    - Supports Python 3.10, 3.11, 3.12, 3.14
  - Automatic virtual environment setup
  - Zero internet required

TROUBLESHOOTING
  - "'python' not recognized" / "Python not found"
    -> Reinstall Python from python.org, check "Add Python to PATH"
  - "Failed to create virtual environment"
    -> You may have the Microsoft Store Python. Uninstall it and
       install Python from https://www.python.org/downloads/
  - Server fails to start
    -> Check if port 5000 is already in use by another application
  - For firewall issues, allow Python through Windows Defender Firewall
