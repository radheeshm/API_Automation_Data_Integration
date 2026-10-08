# AKASH TECH SOLUTIONS
## API Automation & Data Integration

Professional Python/Tkinter desktop application for REST API testing, automation and business data integration.

### Project 03 Features
- Professional AKASH TECH SOLUTIONS green/black UI
- Maximized Windows desktop layout
- AKASH logo in header and Windows taskbar/window icon
- Global status bar with project, task, status, start and end time
- REST API: GET / POST / PUT / DELETE
- Demo/offline API mode for client demonstrations without credentials
- Live API mode using `requests`
- Bearer Token authentication
- API Key authentication
- Record ID support
- GET pagination parameters
- Configurable retry count and timeout
- JSON request body validation
- JSON response viewer
- JSON-to-table transformation
- Nested JSON flattening
- CSV export
- Excel export with API Data and Export Summary sheets
- Request/response logging
- Save/load endpoint configurations
- Background request processing so the UI remains responsive
- Progress indicator and request summary
- Completion dialog
- Help / Settings / About dialogs
- Open output folder
- Built-in sample payload
- Test plan and build instructions

### Run
```powershell
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe main.py
```

PowerShell activation is not required.

### Demo Mode
Demo Mode is enabled by default. It supports GET, POST, PUT and DELETE against a local in-memory dataset.

### Live API
Turn Demo Mode off and provide the endpoint/authentication details. Use a safe test API first.

### Output
Exports are placed in the `output` folder by default. Logs are stored in `logs`.

### Build EXE
```powershell
venv\Scripts\python.exe -m pip install pyinstaller
venv\Scripts\pyinstaller.exe --noconfirm --clean --windowed --name AKASH_API_Automation --icon assets\akash_api.ico main.py
```


## Version 1.1.0 UI improvements
- Improved Transform & Export table visibility with a larger workspace and integrated vertical/horizontal scrollbars.
- Improved table column sizing for readable API data.
- GET and DELETE automatically clear and disable the JSON body because these methods do not require a request body.
- POST and PUT automatically enable the JSON body and provide a demo payload when empty.
- Completion dialog now distinguishes successful 2xx responses from HTTP error responses (4xx/5xx).
- Added a safe fallback when window zooming is unavailable on some Tk environments.
