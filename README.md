# AET: Asana Export Tool
AET (Asana Export Tool) is a tool for quickly scanning and downloading data from Asana.

## Currently Supported:
- Attachments
- Windows EXE

## Planned Support:
- Linux and Mac Compatibility
- Task Data
- Project Data
- Team Data
- Comments/Discussions

And any other suggestions or requests.

PRs are VERY welcome, I will get to them as I have time.

# Usage:
## Get Asana Token:
You will need an Asana API Token to use this app.
1. Open: https://app.asana.com/0/my-apps in a browser.
2. Click `Create New Token` in the bottom left.
3. Copy the token value.
4. SAVE THE TOKEN. You will NOT see it again after copying it, so make sure you save it somewhere.

## Run the App
1. Download the latest release (Windows only for the moment, sorry): https://github.com/GraphicHealer/AsanaExporter/releases/latest
2. Run the EXE (It's all self-contained).
3. Paste in a valid Asana API Token (See above).
4. Press `Connect`
5. If you are connected to more than one workspace, select your workspace from the list.
6. Click the `Scan` button in the upper-right.
7. Wait. It takes it a few minutes to scan, longer if the workspace has alot of tasks or attachments. When the scan is done, it will tell you how many attachments you have.
8. Click the `Browse` button and select an export directory. Make sure it's empty.
9.  Once the scan is complete and the export directory is selected, click the `Start Download` button to download all the attachement files.
10. Wait. This takes a while. All of your attachments are being downloaded into folders under project name.
11. Once done, you can click the `Open Folder` button to lauch file explorer and view the downloaded attachments!

# Building from Source:
This assumes Python is installed and running on your system.

## Setup
1. Clone this github:
    ```
    git clone https://github.com/GraphicHealer/AsanaExporter.git
    ```
2. Create a VENV:
    ```
    python -m venv .
    ```
3. Install Requirements:
    ```
    python -m pip install requirements.txt
    ```

## Build
To build the EXE, run the following in a terminal opened in the Python VENV:
```
pyinstaller AsanaExport.spec
```
the result will be saved in the `dist` folder as `AsanaExport.exe`.
