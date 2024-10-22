######################
# AsanaExport.py
# v1.0.0
# By: GrapicHealer
######################



## INITIALIZATION ##

# Import Libraries
from pathvalidate import sanitize_filename
from pathvalidate import sanitize_filepath
from asana.rest import ApiException
from threading import Thread
import FreeSimpleGUI as sg
import urllib.request
import pathlib
import asana
import time
import ast
import os
import re

# Initalize Global Variables
log = []
Status = ''
attachments = []
folders = []
noUrl = []
debug_view = False

# Get Icon
image = os.path.join(os.path.dirname(__file__), 'icons\\asana.ico')



## FUNCTIONS ##

# Function for Logging
def logger(win, string, label=False):
    if (label == True):
        win['LABEL'].update(str(string).strip())
    log.append(str(string))
    win['LOG'].update('\n'.join(log))
    win.refresh()

# Function for Asana Workspace API
def getWorkspaces(win, api_instance):
    try:
        api_response = api_instance.get_workspaces({})
        return list(api_response)
    except ApiException as e:
        logger(win, "Exception when calling UsersApi->get_workspaces: %s\n" % e)

# Function for Asana Team API
def getTeams(win, api_instance, workspace):
    try:
        api_response = api_instance.get_teams_for_workspace(workspace['gid'], {})
        return list(api_response)
    except ApiException as e:
        logger(win, "Exception when calling TeamsApi->get_teams_for_workspace: %s\n" % e)

# Function for Asana Project API
def getProjects(win, api_instance, data):
    try:
        api_response = api_instance.get_projects({'team': data['gid']})
        return list(api_response)
    except ApiException as e:
        logger(win, "Exception when calling ProjectsApi->get_projects: %s\n" % e)

# Function for Asana Task API
def getTasks(win, api_instance, data):
    try:
        api_response = api_instance.get_tasks({'project': data['gid']})
        return list(api_response)
    except ApiException as e:
        logger(win, "Exception when calling TasksApi->get_tasks: %s\n" % e)

# Function for Asana Attachment API
def getAttachments(win, api_instance, data):
    try:
        api_response = api_instance.get_attachments_for_object(data['gid'], {'opt_fields': 'download_url,name'})
        return list(api_response)
    except ApiException as e:
        logger(win, "Exception when calling AttachmentsApi->get_attachments_for_object: %s\n" % e)

# Function to Find Attachment Files
def getFiles(win, task_api, attachment_api, project):
    proj_name = re.sub(' {2,}',' ',sanitize_filename(''.join(i for i in project['name'] if ord(i)<128)))
    folderCount = 0
    folderEnd = ''
    for folder in folders:
        if proj_name in folder:
            folderEnd = ' (' + str(folderCount) + ')'
            folderCount += 1
    proj_name = proj_name + folderEnd
    folders.append(proj_name)
    attachment_list = getAttachments(win, attachment_api, project)
    tasks = getTasks(win, task_api, project)
    length = len(tasks)
    count = 0
    win['PROBAR1'].update(max=length, current_count=0)
    win['PROLAB1'].update("'"+proj_name+"' Tasks")
    logger(win, "   Parsing Tasks...", label=True)
    for task in tasks:
        count += 1
        win['PROBAR1'].update(current_count=count)
        win['PROCOUNT1'].update(str(count) + "/" + str(length))
        logger(win, "    - '"+task['name']+"'")
        attachment_list.extend(getAttachments(win, attachment_api, task))

    logger(win, "   Parsing Attachments...", label=True)
    output = []
    files = []
    length = len(attachment_list)
    count = 0
    win['PROBAR1'].update(max=length, current_count=0)
    win['PROLAB1'].update("'"+proj_name+"' Attachments")
    for attachment in attachment_list:
        count += 1
        fileName = re.sub(' {2,}',' ',sanitize_filename(''.join(i for i in attachment['name'] if ord(i)<128)))
        fileCount = 0
        fileEnd = ''
        for file in files:
            if fileName in file:
                fileEnd = ' (' + str(fileCount) + ')'
                fileCount += 1
        fileName = fileName + fileEnd
        files.append(fileName)
        win['PROBAR1'].update(current_count=count)
        win['PROCOUNT1'].update(str(count) + "/" + str(length))
        logger(win, attachment['download_url']) if debug_view else None
        if (attachment['download_url'] == None):
            logger(win, "-------- Attachment has no download URL. --------") if debug_view else None
            noUrl.append(attachment)
            continue
        output.append({
            'gid': attachment['gid'],
            'url': attachment['download_url'],
            'folder': proj_name,
            'name': fileName
        })
        logger(win, "    - "+str(output[-1]['name'])+" ("+str(count)+"/"+str(len(attachment_list))+")") if debug_view else None
    return output

# Function to Download a File
def downloadFile(win, url, filepath):
    filepath = sanitize_filepath(filepath)

    win['PROBAR1'].update(max=100,current_count=0)
    win['PROLAB1'].update('Downloading: '+os.path.basename(filepath))
    win.refresh()
    try:
        pathlib.Path(os.path.dirname(filepath)).mkdir(parents=True, exist_ok=True)
        urllib.request.urlretrieve(url, filepath, progress)
    except Exception as e:
        logger(win, 'URL: '+str(url)+'\n'+'Error: '+str(e)) if debug_view else None

# Function to Track Download Progress
def progress(count, block_size, total_size):
    global start_time
    global window
    if (count == 0):
        start_time = time.time()
        return
    duration = time.time() - start_time
    progress_size = int(count * block_size) / (1024 * 1024)
    speed = int(progress_size / (1024 * (duration + 1)))
    percent = min(int(count * block_size * 100 / total_size), 100)

    window['PROBAR1'].update(current_count=percent)
    window['PROCOUNT1'].update(str(percent) + "%")
    window.refresh()

# Function to Scan Project for Attachments
def window_scan(win, values, teams_api, projects_api, attachments_api, tasks_api, workspace):
    global Status
    global attachments
    Status = 'Scanning'
    attachments = []
    teams = getTeams(win, teams_api, workspace)
    length = len(teams)
    logger(win, str(length) + ' Teams Found.', label=True)
    projects = []
    logger(win, 'Retriving Teams...', label=True)
    win['PROBAR1'].update(max=length, current_count=0)
    win['PROLAB1'].update('Teams')
    win['PROG'].update(visible=True)
    win['QUIT'].update('Cancel')
    win['Scan'].update(disabled=True)
    win.refresh()
    count = 0
    for team in teams:
        if Status == 'Cancel':
            exit()
            break
        count += 1
        win['PROBAR1'].update(current_count=count)
        win['PROCOUNT1'].update(str(count) + "/" + str(length))
        logger(win, ' - '+str(team['name']))
        logger(win, ' - '+str({x: team[x] for x in team if x not in ["gid"]})) if debug_view else None
        projects.extend(getProjects(win, projects_api, team))
    logger(win, str(len(teams)) + ' Teams Scanned.\n')
    win['PROG'].update(visible=False)
    win['PROG2'].update(visible=True)
    win.refresh()

    logger(win, 'Retriving Projects...', label=True)
    count = 0
    length = len(projects)
    win['PROBAR2'].update(max=length, current_count=0)
    win['PROLAB2'].update('Projects')
    win['PROG'].update(visible=True)
    win.refresh()
    for project in projects:
        if Status == 'Cancel':
            exit()
            break
        count += 1
        win['PROBAR2'].update(current_count=count)
        win['PROCOUNT2'].update(str(count) + "/" + str(length))
        logger(win, ' - ' + project['name'])
        logger(win, ' - '+str({x: project[x] for x in project if x not in ["gid"]})) if debug_view else None
        attachments.extend(getFiles(win, tasks_api, attachments_api, project))
    logger(win, str(len(projects)) + ' Projects Scanned.\n', label=True)

    logger(win, "Scan Complete! "+str(len(attachments)) + ' Attachments Found.', label=True)
    logger(win, 'Attachments with NO Download URL:\n'+'\n'.join(noUrl)) if debug_view else None
    win['PROG'].update(visible=False)
    win['EXPORT'].update(visible=True)
    win['QUIT'].update('Quit')
    win['EXPORT_ATTACH'].update(disabled=False)
    Status = 'Scanned'
    win.refresh()

    if ((values['Browse'] != 'Select an Export Directory...') and (values['Browse'] != '') and (attachments != [])):
        win['DOWNLOAD'].update(disabled=False)
    else:
        logger(win, "Click 'Browse' to Select an Export Directory.", label=True)

# Function to Download Files
def window_download(win, values):
    global Status
    global attachments
    Status = 'Downloading'
    logger(win, 'Downloading Attachments...', label=True)
    logger(win, '')
    export_dir = values['Browse'].replace('/', '\\')
    length = len(attachments)
    count = 0
    win['PROBAR2'].update(max=length, current_count=0)
    win['PROLAB2'].update('Attachments')
    win['PROG'].update(visible=True)
    win['PROG2'].update(visible=True)
    win['QUIT'].update('Cancel')
    win['DOWNLOAD'].update(disabled=True)
    win['Browse'].update(disabled=True)
    win.refresh()
    for file in attachments:
        if Status == 'Cancel':
            exit()
            break
        count += 1
        filepath = os.path.join(export_dir, file['folder'], file['name'])
        win['PROBAR2'].update(current_count=count)
        win['PROCOUNT2'].update(str(count) + "/" + str(length))
        logger(win, "Downloading '"+file['name']+"' to '"+filepath+"' ("+str(count)+"/"+str(length)+")")
        downloadFile(win, file['url'], filepath)
    logger(win, "Download Complete! Files Downloaded to '"+export_dir+"'", label=True)
    win['PROG'].update(visible=False)
    win['DOWNLOAD'].update(disabled=False)
    win['DOWNLOAD'].update('Open Folder')
    win['QUIT'].update('Quit')
    Status = 'Done'
    win.refresh()

# Function to Enable/Disable GUI Sections
def collapse(layout, key, visible):
    return sg.pin(sg.Column(layout, key=key, visible=visible, expand_x=True, expand_y=True, pad=(0), element_justification='center'), expand_x=True, expand_y=True)



## Main Program ##

# GUI Layout
api_layout = [
    [sg.Text('Enter API Key:', right_click_menu=['', ['Debug View']]), sg.Push()],
    [sg.Input(expand_x=True, key='TOKEN'), sg.Button('Connect')]
]

workspace_layout = [
    [sg.Text('Pick a Workspace:'), sg.Push()],
    [sg.Listbox(values=[], justification='center', enable_events=True, size=(50, 10), expand_x=True, no_scrollbar=True, key='BOX', select_mode='LISTBOX_SELECT_MODE_SINGLE')]
]

scanner_layout = [
    [sg.Text("Click 'Scan' to Find Attachments", key='LABEL'), sg.Push(), sg.Button('Scan'), sg.Button('Export Log', visible=False, key='EXPORT_LOG'), sg.Button('Export Attachments', visible=False, key='EXPORT_ATTACH')],
    [sg.Multiline(disabled=True, autoscroll=True, autoscroll_only_at_bottom=True, size=(50, 10), expand_x=True, key='LOG')]
]

progress_layout = [
    [sg.Text('Item', key='PROLAB1'), sg.Push(), sg.Text('', key='PROCOUNT1')],
    [sg.ProgressBar(100, key='PROBAR1', orientation='h', size=(20, 20), expand_x=True)],
    [collapse([
        [sg.Text('Total', key='PROLAB2'), sg.Push(), sg.Text('', key='PROCOUNT2')],
        [sg.ProgressBar(100, key='PROBAR2', orientation='h', size=(20, 20), expand_x=True)]
    ], 'PROG2', False)]
]

export_layout = [
    [sg.Input('Select an Export Directory...', disabled=True, enable_events=True, key='Browse', expand_x=True), sg.FolderBrowse()],
    [sg.Button('Start Download', key='DOWNLOAD', disabled=True)]
]

layout = [
    [collapse(api_layout, 'API', True)],
    [collapse(workspace_layout, 'WORK', False)],
    [collapse(scanner_layout, 'SCAN', False)],
    [collapse(progress_layout, 'PROG', False)],
    [collapse(export_layout, 'EXPORT', False)],
    [collapse([[sg.Button('Quit', key='QUIT')]], 'EXIT', False)]
]

# Initialize GUI
window = sg.Window('Asana Attachment Exporter', layout, enable_close_attempted_event=True, finalize=True, icon=image)

# Main Loop
while True:
    event, values = window.read()
    # print(event, values)

    if ((event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT) or (event == 'QUIT')):
        match Status:
            case 'Scanning':
                if (sg.popup_yes_no('There is an active scan, Are you sure you want to cancel?', title='Exit') == 'Yes'):
                    Status = 'Cancel'
                    break
            case 'Scanned':
                if (sg.popup_yes_no('You will lose your current scan, Are you sure you want to quit?', title='Exit') == 'Yes'):
                    Status = 'Cancel'
                    break
            case 'Downloading':
                if (sg.popup_yes_no('A download is in progress, Are you sure you want to cancel?', title='Exit') == 'Yes'):
                    Status = 'Cancel'
                    break
            case _:
                break

    if (event == 'Debug View'):
        debug_view = not debug_view
        if (debug_view == True):
            logger(window, 'Debug View Enabled', label=True)
            window['LOG'].set_size(size=(100, 20))
            window['EXPORT_LOG'].update(visible=True)
            window['EXPORT_ATTACH'].update(visible=True)
        else:
            window['LOG'].set_size(size=(50, 10))
            window['EXPORT_LOG'].update(visible=False)
            window['EXPORT_ATTACH'].update(visible=False)

    if (event == 'Connect'):
        if values['TOKEN'] == '':
            sg.popup_ok('Please Enter an API Key.', title='Error')
            continue
        configuration = asana.Configuration()
        configuration.access_token = values['TOKEN']
        api_client = asana.ApiClient(configuration)

        workspaces_api = asana.WorkspacesApi(api_client)
        teams_api = asana.TeamsApi(api_client)
        projects_api = asana.ProjectsApi(api_client)
        tasks_api = asana.TasksApi(api_client)
        attachments_api = asana.AttachmentsApi(api_client)

        logger(window, 'Scanning Workspaces...', label=True) if debug_view else None
        workspaces = getWorkspaces(window, workspaces_api)
        spaces = []

        for space in workspaces:
            logger(window, ' - '+str({x: space[x] for x in space if x not in ["gid"]})) if debug_view else None
            spaces.append(space['name'])

        window['BOX'].update(spaces)
        window['API'].update(visible=False)
        window['WORK'].update(visible=True)
        window['EXIT'].update(visible=True)

        if (len(spaces) == 1):
            window.write_event_value('BOX', spaces)

    if ((event == 'BOX') and (values['BOX'] != [])):
        workspace = next(item for item in workspaces if item['name'] == values['BOX'][0])
        window['WORK'].update(visible=False)
        logger(window, "Connected to '" + workspace['name'] + "' Asana Workspace.", label=True)
        window['SCAN'].update(visible=True)
        logger(window, "Click 'Scan' to Get Attachments", label=True)

    if (event == 'Scan'):
        if ((debug_view == True) and (os.path.exists('Attachments.txt') == True)):
            logger(window, ' Attachments.txt Found, Importing...', label=True)
            f = open('Attachments.txt', 'r')
            for line in f:
                attachments.append(ast.literal_eval(line.rstrip()))
            f.close()
            window['EXPORT'].update(visible=True)
            window['QUIT'].update('Quit')
            Status = 'Scanned'

            if ((values['Browse'] != 'Select an Export Directory...') and (values['Browse'] != '') and (attachments != [])):
                window['DOWNLOAD'].update(disabled=False)
                logger(window, "Click 'Start Download' to Download All Attachments.", label=True)
            else:
                logger(window, "Click 'Browse' to Select an Export Directory.", label=True)
        else:
            scanThread = Thread(target=window_scan, args=(window, values, teams_api, projects_api, attachments_api, tasks_api, workspace), daemon=True).start()


    if (event == 'Browse'):
        if ((values['Browse'] == '') or (values['Browse'] == 'Select an Export Directory...')):
            logger(window, 'Missing or Invalid Directory.', label=True)
            window['DOWNLOAD'].update(disabled=True)
        elif (attachments != []):
            logger(window, 'Directory Selected: '+values['Browse'], label=True)
            logger(window, "Click 'Start Download' to Download All Attachments.", label=True)
            window['DOWNLOAD'].update(disabled=False)

    if (event == 'EXPORT_LOG'):
        f = open('Log_Export.txt', 'w')
        for line in log:
            f.write(str(line)+'\n')
        f.close()

    if (event == 'EXPORT_ATTACH'):
        f = open('Attachments.txt', 'w')
        for line in attachments:
            f.write(str(line)+'\n')
        f.close()

    if (event == 'DOWNLOAD'):
        if (Status == 'Done'):
            os.startfile(values['Browse'])
        elif (Status == 'Scanned'):
            downloadThread = Thread(target=window_download, args=(window, values), daemon=True).start()
            # window_download(window, values)

# Close GUI
window.close()
