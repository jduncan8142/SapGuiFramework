# Getting Started
You will need the below tooling installed and configured to follow along with these tutorials.

- [SAP GUI for Windows](https://help.sap.com/docs/sap_gui_for_windows)
- [Python version 3.11+ 64 bit](https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe)
- [Scripting Tracker by Stefan Schnell](https://tracker.stschnell.de/)
- [VS Code](https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user)
    - Plug-ins:
     - [Python Extension Pack](https://marketplace.visualstudio.com/items?itemName=donjayamanne.python-extension-pack)
     - [Jupyter Notebook](https://marketplace.visualstudio.com/items?itemName=ms-toolsai.jupyter)
     - [Json](https://marketplace.visualstudio.com/items?itemName=ZainChen.json)
- [SAP GUI Framework](https://github.com/jduncan8142/SapGuiFramework.git)

## Installation Steps for Windows
### Installing SAP GUI for Windows
It is assumed that SAP GUI for Windows software is already installed on your system and the details are not covered here. If more information on installation is needed it can be found [here](https://help.sap.com/docs/sap_gui_for_windows/1ebe3120fd734f67afc57b979c3e2d46/78cb9f653b1c465c9f1b7009c515c94e.html). 

### Installing Python
Also, it is assumed you can install Python on your own but if additional information is required it can be found [here](https://docs.python.org/3/using/windows.html) on the Using Python on Windows documentation page. 

*Note: While installing Python be sure to select the Add Python to Path checkbox on the first installer screen.*

### Installing Scripting Tracker


### Installing VS Code


### Installing VS Code Plug-ins


### Creating a Python Virtual Environment
1. Create virtual environment: `pipenv --python 3.11`
2. Enter into the new virtual environment: `pipenv shell`

## Installing SapGuiFramework
```powershell
pipenv install 'SapGuiFramework @ git+https://github.com/jduncan8142/SapGuiFramework.git@main'
```

### Updating SapGuiFramework
```powershell
pipenv uninstall sapguiframework; pipenv install 'SapGuiFramework @ git+https://github.com/jduncan8142/SapGuiFramework.git@main'
```

### .env Files
If you need to provide sensitive data such as usernames, passwords, or private URLs to your test scripts, this can be accomplished using a .env file. 

Simply create a new file in the root of your test directory named `.env`

*Note: Make sure to include the dot (.) at the beginning of the file name. The file must be named .env exactly. No other naming is allowed.*

SapGuiFramework will automatically read this .env file from the root of the test directory when you create a session instance. You then can refer to the values defined within .env like below: 

```python
from Core.Framework import Session

sap = Session()

sap.get_env("PASSWORD")
```

The syntax of `.env` file should be as below:

```bash
# Comment start with # and continue until the end of the line.
USERNAME=my_username
PASSWORD=my_secret_password
DOMAIN=example.org
ROOT_URL=${DOMAIN}/app
```

If you would like to understand how this process is working you can find more information [here](https://pypi.org/project/python-dotenv/).