from pathvalidate import sanitize_filename
from pathvalidate import sanitize_filepath
from asana.rest import ApiException
from threading import Thread
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
import urllib.request
import pathlib
import asana
import os
import re

class AsanaAPI:
    def __init__(self, api_token):
        configuration = asana.Configuration()
        configuration.access_token = api_token
        api_client = asana.ApiClient(configuration)

        self.workspaces_api = asana.WorkspacesApi(api_client)
        self.teams_api = asana.TeamsApi(api_client)
        self.projects_api = asana.ProjectsApi(api_client)
        self.tasks_api = asana.TasksApi(api_client)
        self.attachments_api = asana.AttachmentsApi(api_client)

    class Object:
        raw = {}
        def __init__(self, data, **kwargs):
            self.raw = data | kwargs
            for key, value in self.raw.items():
                setattr(self, key, value)

    def getWorkspaces(self, logger=None, parent=None):
        try:
            api_response = list(self.workspaces_api.get_workspaces({}))

            for i in range(len(api_response)):
                api_response[i] = self.Object(api_response[i])
            return api_response
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling WorkspaceApi->get_workspaces: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling WorkspaceApi->get_workspaces: %s\n" % e)

    # Function for Asana Team API
    def getTeams(self, workspace_object, logger=None, parent=None):
        try:
            api_response = list(self.teams_api.get_teams_for_workspace(workspace_object.gid, {}))
            for i in range(len(api_response)):
                api_response[i] = self.Object(api_response[i], workspace=workspace_object)
            return api_response
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling TeamsApi->get_teams_for_workspace: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling TeamsApi->get_teams_for_workspace: %s\n" % e)

    # Function for Asana Project API
    def getProjects(self, team_object, logger=None, parent=None):
        try:
            api_response = list(self.projects_api.get_projects({'team': team_object.gid}))
            for i in range(len(api_response)):
                api_response[i] = self.Object(api_response[i], team=team_object, workspace=team_object.workspace)
            return api_response
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling ProjectsApi->get_projects: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling ProjectsApi->get_projects: %s\n" % e)

    # Function for Asana Task API
    def getTasks(self, project_object, logger=None, parent=None):
        try:
            api_response = list(self.tasks_api.get_tasks({'project': project_object.gid}))
            for i in range(len(api_response)):
                api_response[i] = self.Object(api_response[i], project=project_object, team=project_object.team, workspace=project_object.workspace)
            return api_response
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling TasksApi->get_tasks: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling TasksApi->get_tasks: %s\n" % e)

    # Function for Asana Attachment API
    def getAttachments(self, input_object, logger=None, parent=None):
        if input_object.resource_type == 'task':
            var_input = {'task': input_object, 'project': input_object.project}
        elif input_object.resource_type == 'project':
            var_input = {'project': input_object}
        try:
            api_response = list(self.attachments_api.get_attachments_for_object(input_object.gid, {'opt_fields': 'download_url,name'}))
            for i in range(len(api_response)):
                api_response[i] = self.Object(api_response[i], **var_input, team=input_object.team, workspace=input_object.workspace, resource_type=input_object.resource_type, path='')
            return api_response
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling AttachmentsApi->get_attachments_for_object: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling AttachmentsApi->get_attachments_for_object: %s\n" % e)

    def getAttachmentURL(self, input_object, logger=None, parent=None):
        try:
            api_response = self.attachments_api.get_attachment(input_object.gid, {'opt_fields': 'download_url'})
            return api_response['download_url']
        except ApiException as e:
            if logger is not None:
                logger("Exception when calling AttachmentsApi->get_attachment: %s\n" % e, parent=parent, popen=True)
            else:
                print("Exception when calling AttachmentsApi->get_attachment: %s\n" % e)

class FancyProgressBar:
    def __init__(self, layout):
        self.prog_layout = ttk.Frame(layout)

        self.label = StringVar()
        self.label.set('')
        ttk.Label(self.prog_layout, textvariable=self.label).grid(column=0, row=0, sticky=(N,W), padx=6)

        self.count = StringVar()
        self.count.set('0/0')
        ttk.Label(self.prog_layout, textvariable=self.count).grid(column=2, row=0, sticky=(N,E), padx=6)

        self.bar_var = IntVar()
        self.bar_max = IntVar(value=100)
        self.bar = ttk.Progressbar(self.prog_layout, orient=HORIZONTAL, mode='determinate', variable=self.bar_var, maximum=self.bar_max.get())
        self.bar.grid(column=0, row=1, columnspan=3, sticky=(W,E), padx=6)

    def grid(self, *args, **kwargs):
        self.prog_layout.grid(*args, **kwargs)
        self.prog_layout.columnconfigure(list(range(3)), weight=1)

    def setup(self, maximum, label, visible=True, percent=False):
        self.bar_var.set(0)
        self.label.set(label)
        if (percent == True):
            self.count.set('0%')
        else:
            self.count.set('0/' + str(maximum))
        self.bar.config(maximum=maximum)
        self.bar_max.set(maximum)
        self.set_visible(visible)
        self.percent = percent
        return self

    def set(self, value):
        if (self.percent == True):
            self.count.set(str(value) + '%')
        else:
            self.count.set(str(value) + '/' + str(self.bar_max.get()))
        self.bar_var.set(value)

    def get(self):
        return self.bar_var.get()

    def set_visible(self, visible):
        self.visible = visible
        if (visible == True):
            self.prog_layout.grid()
        else:
            self.prog_layout.grid_remove()

class AsanaExport:

    def __init__(self, root):

        self.debug_view = False
        self.status = 'Ready'
        self.projects = []
        self.attachments = []
        self.tasks = []

        image = os.path.join(os.path.dirname(__file__), 'icons/asana.ico')

        self.root = root
        self.root.title("Asana Export")
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.iconbitmap(image)

        mainframe = ttk.Frame(self.root, padding="3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N,W,E,S))
        mainframe.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # API layout
        self.api_layout = ttk.Frame(mainframe)
        self.api_layout.grid(column=0, row=0, sticky=(N,W,E,S), pady=0, padx=6)
        self.api_layout.columnconfigure(list(range(2)), weight=1)

        ttk.Label(self.api_layout, text="API Key:").grid(column=0, row=0, sticky=(N,W), padx=6, pady=0)

        self.api_token = StringVar()
        token_entry = ttk.Entry(self.api_layout, width=50, textvariable=self.api_token)
        token_entry.grid(column=0, row=1, columnspan=2, sticky=(N,W,E,S), pady=0, padx=6)
        token_entry.focus()

        ttk.Button(self.api_layout, text="Connect", command=self.connect_api).grid(column=2, row=1, sticky=(N,E,S), pady=0, padx=6)

        # Workspace layout
        self.workspace_layout = ttk.Frame(mainframe)
        self.workspace_layout.grid(column=0, row=0, sticky=(N,W,E,S), pady=0, padx=6)
        self.workspace_layout.columnconfigure(list(range(2)), weight=1)
        self.hide(self.workspace_layout)

        ttk.Label(self.workspace_layout, text="Select a Workspace:").grid(column=0, row=0, sticky=(N,W), padx=6, pady=0)

        self.workspace_var = StringVar()
        self.workspace_combo = ttk.Combobox(self.workspace_layout, width=50, textvariable=self.workspace_var)
        self.workspace_combo.grid(column=0, row=1, columnspan=2, sticky=(N,W,E,S), padx=6, pady=0)
        self.workspace_combo.state(["readonly"])

        ttk.Button(self.workspace_layout, text="Select", command=self.select_workspace).grid(column=2, row=1, sticky=(N,E,S), padx=6, pady=0)

        # Scan layout
        self.scan_layout = ttk.Frame(mainframe)
        self.scan_layout.grid(column=0, row=0, sticky=(N,W,E,S), padx=6, pady=0)
        self.scan_layout.columnconfigure(list(range(2)), weight=1)
        self.scan_layout.rowconfigure(1, weight=1)
        self.hide(self.scan_layout)

        self.label_var = StringVar()
        self.label_var.set("Click 'Scan' to Get Attachments")
        self.label = ttk.Label(self.scan_layout, textvariable=self.label_var)
        self.label.grid(column=0, row=0, sticky=(N,W,S), padx=6, pady=0)

        self.scan_button = ttk.Button(self.scan_layout, text="Scan", command=lambda: Thread(target=self.scan_workspace, daemon=True).start())
        self.scan_button.grid(column=2, row=0, sticky=(N,E,S), padx=6, pady=0)

        self.log_tree = ttk.Treeview(self.scan_layout, height=20, selectmode='none', show='tree')
        self.log_tree.grid(column=0, row=1, columnspan=3, sticky=(N,W,E,S), padx=6, pady=0)
        self.log_tree.column("#0", width=500, stretch=True)

        # Progress Layout
        self.progress_layout = ttk.Frame(mainframe)
        self.progress_layout.grid(column=0, row=1, sticky=(N,W,E,S), padx=6, pady=0)
        self.progress_layout.columnconfigure(0, weight=1)
        self.hide(self.progress_layout)

        self.total_progress = FancyProgressBar(self.progress_layout)
        self.total_progress.grid(column=0, row=0, columnspan=3, sticky=(N,W,E,S), padx=6, pady=0)
        self.total_progress.setup(100, 'Item')

        self.item_progress = FancyProgressBar(self.progress_layout)
        self.item_progress.grid(column=0, row=1, columnspan=3, sticky=(N,W,E,S), padx=6, pady=0)
        self.item_progress.setup(100, 'Total', visible=False)

        # Export Layout
        self.export_layout = ttk.Frame(mainframe)
        self.export_layout.grid(column=0, row=2, sticky=(N,W,E,S), padx=6, pady=0)
        self.export_layout.columnconfigure(list(range(2)), weight=1)
        self.hide(self.export_layout)

        self.export_path = StringVar()
        self.browse_input = ttk.Entry(self.export_layout, width=50, textvariable=self.export_path, state='disabled')
        self.browse_input.grid(column=0, row=0, columnspan=2, sticky=(N,W,E,S), padx=6, pady=0)
        self.browse_input.focus()

        self.browse_button = ttk.Button(self.export_layout, text="Browse", command=self.browse_folder)
        self.browse_button.grid(column=2, row=0, sticky=(N,E,S), padx=6, pady=0)

        self.download_button = ttk.Button(self.export_layout, text="Start Download", command=lambda: Thread(target=self.download_attachments, daemon=True).start(), state='disabled')
        self.download_button.grid(column=2, row=1, sticky=(N,E,S,W), padx=6, pady=0)

        self.open_folder_button = ttk.Button(self.export_layout, text="Open Folder", command=self.open_folder)
        self.open_folder_button.grid(column=2, row=2, sticky=(N,E,S,W), padx=6, pady=0)
        self.hide(self.open_folder_button)

    def logger(self, string, id=None, parent='', label=False, open=False, popen=False):
        string = str(string).strip()
        if (label == True):
            self.label_var.set(string)
        if (self.debug_view == True):
            open = True
        if (popen == True):
            self.log_tree.item(parent, open=True)
        self.log_tree.insert(parent=parent, index='end', iid=id, text=string, open=open)
        self.log_tree.yview_moveto(1)

    def connect_api(self):
        if (self.api_token == ''):
            print('Please Enter an API Key')
            return

        self.API = AsanaAPI(self.api_token.get())

        self.logger('Finding Workspaces...', id='Workspaces', label=True)
        self.workspaces = self.API.getWorkspaces(logger=self.logger, parent='Workspaces')
        spaces = []

        for space in self.workspaces:
            self.logger(str(space.raw), parent='Workspaces')
            spaces.append(space.name)

        self.workspace_combo['values'] = spaces
        self.workspace_var.set(spaces[0])
        self.hide(self.api_layout)

        if (len(spaces) == 1):
            self.select_workspace()
        else:
            self.show(self.workspace_layout)
            self.label_var.set('Select Workspace:')

    def select_workspace(self):
        self.workspace = next(item for item in self.workspaces if item.name == self.workspace_var.get())
        self.logger("Connected to '" + self.workspace.name + "' Asana Workspace.", label=True)
        self.logger("Click 'Scan' to Get Attachments", label=True)
        self.hide(self.workspace_layout)
        self.show(self.scan_layout)

    def scan_workspace(self):
        self.scan_button['state'] = 'disabled'
        self.status = 'Scanning'
        self.logger('Scanning Workspace...', id='Workspace', label=True)
        self.teams = self.API.getTeams(self.workspace, logger=self.logger, parent='Workspace')
        self.teams = self.format_object_names(self.teams)
        self.logger("Found '" + str(len(self.teams)) + "' Teams.", label=True)

        self.scan_teams()
        self.scan_projects()
        self.scan_tasks()
        self.parse_attachments()

        self.hide(self.progress_layout)
        self.show(self.export_layout)
        self.status = 'Scanned'
        self.logger("Click 'Browse' to Select an Export Directory.", label=True)

    def scan_teams(self):
        self.total_progress.setup(len(self.teams), 'Teams')
        self.show(self.progress_layout)
        self.logger('Scanning Teams...', id='Teams', label=True)
        for i in range(len(self.teams)):
            if self.status == 'Cancel':
                exit()
                break
            self.total_progress.set(i+1)
            self.logger(self.teams[i].name, parent='Teams', id=self.teams[i].gid)
            self.logger(self.teams[i].raw, parent=self.teams[i].gid)
            self.projects.extend(self.API.getProjects(self.teams[i], logger=self.logger, parent=self.teams[i].gid))
        self.projects = list(filter(None, self.projects))
        self.projects = self.format_object_names(self.projects)
        self.logger('Found ' + str(len(self.projects)) + ' Projects.', label=True)

    def scan_projects(self):
        self.total_progress.setup(len(self.projects), 'Projects')
        self.logger('Scanning Projects...', id='Projects', label=True)
        for i in range(len(self.projects)):
            if self.status == 'Cancel':
                exit()
                break
            self.total_progress.set(i+1)
            self.logger(self.projects[i].name, parent='Projects', id=self.projects[i].gid)
            self.logger(self.projects[i].raw, parent=self.projects[i].gid)
            self.attachments.extend(self.API.getAttachments(self.projects[i], logger=self.logger, parent=self.projects[i].gid))
            self.tasks.extend(self.API.getTasks(self.projects[i], logger=self.logger, parent=self.projects[i].gid))
        self.tasks = self.format_object_names(self.tasks)
        self.logger('Found ' + str(len(self.tasks)) + ' Tasks.', label=True)

    def scan_tasks(self):
        self.total_progress.setup(len(self.tasks), 'Tasks')
        self.logger('Scanning Tasks...', id='Tasks', label=True)
        for i in range(len(self.tasks)):
            if self.status == 'Cancel':
                exit()
                break
            self.total_progress.set(i+1)
            self.logger(self.tasks[i].name, parent='Tasks', id=self.tasks[i].gid)
            self.logger(self.tasks[i].raw, parent=self.tasks[i].gid)
            self.attachments.extend(self.API.getAttachments(self.tasks[i], logger=self.logger, parent=self.tasks[i].gid))
        self.attachments = self.format_object_names(self.attachments, attachments=True)
        self.logger('Found ' + str(len(self.attachments)) + ' Attachments.', label=True)

    def parse_attachments(self):
        self.total_progress.setup(len(self.attachments), 'Attachments')
        self.logger('Parsing Attachments...', id='Attachments', label=True)
        for i in range(len(self.attachments)):
            path = []
            if self.status == 'Cancel':
                exit()
                break
            self.total_progress.set(i+1)
            self.logger(self.attachments[i].name, parent='Attachments', id=self.attachments[i].gid)
            self.logger(self.attachments[i].raw, parent=self.attachments[i].gid)
            path.append(self.attachments[i].workspace.name)
            path.append(self.attachments[i].team.name)
            path.append(self.attachments[i].project.name)
            path.append(self.attachments[i].task.name) if self.attachments[i].resource_type == 'task' else None
            path.append(self.attachments[i].name)
            self.attachments[i].path = '/'.join(path)
        self.logger('Parsed ' + str(len(self.attachments)) + ' Attachments.', label=True)

    def format_object_names(self, object_list, attachments=False):
        object_list = [o for o in object_list if o.name]
        if attachments:
            object_list = [o for o in object_list if o.download_url]
        object_names = [re.sub(' {2,}',' ',sanitize_filename(''.join(j for j in str(o.name) if ord(j)<128))).rstrip() for o in object_list if o.name]
        for i in range(len(object_list)):
            name = object_names[i]
            count = object_names.count(name)
            if count > 1:
                if attachments:
                    ext = name.split('.')[-1]
                    name = name.rstrip('.'+ext) + '(' + str(count-1) + ').' + ext
                else:
                    name = name + '(' + str(count-1) + ')'
            object_list[i].name = name
            object_names[i] = name
        return object_list

    def browse_folder(self):
        filename = filedialog.askdirectory()
        self.export_path.set(filename)
        self.download_button['state'] = 'normal'

    def open_folder(self):
        os.startfile(self.export_path.get())

    def download_attachments(self):
        self.download_button['state'] = 'disabled'
        self.browse_button['state'] = 'disabled'
        self.status = 'Downloading'
        self.logger('Downloading Attachments...', id='Downloads', label=True)
        self.item_progress.set_visible(True)
        self.total_progress.setup(len(self.attachments), 'Downloads')
        self.show(self.progress_layout)
        for i in range(len(self.attachments)):
            if self.status == 'Cancel':
                exit()
                break
            self.total_progress.set(i+1)
            self.logger(self.attachments[i].name, parent='Downloads', id='d_'+str(self.attachments[i].gid))
            self.logger(self.attachments[i].raw, parent='d_'+str(self.attachments[i].gid))
            self.download_file(self.API.getAttachmentURL(self.attachments[i], logger=self.logger, parent='d_'+str(self.attachments[i].gid)), self.attachments[i].path, parent='d_'+str(self.attachments[i].gid))
        self.logger('Download Complete.', label=True)
        self.total_progress.setup(len(self.attachments), 'Downloads').set(len(self.attachments))
        self.item_progress.setup(100, 'Download Complete!', percent=True).set(100)
        self.status = 'Done'
        self.hide(self.download_button)
        self.show(self.open_folder_button)

    def download_file(self, url, path, parent=None):
        download_path = sanitize_filepath(os.path.join(self.export_path.get(), path))

        self.item_progress.setup(100, 'Downloading: ' + os.path.basename(download_path), percent=True)

        try:
            pathlib.Path(os.path.dirname(download_path)).mkdir(parents=True, exist_ok=True)
            urllib.request.urlretrieve(url, download_path, self.progress)
        except Exception as e:
            self.logger('Error: '+str(e), parent=parent, popen=True)

    def progress(self, block_num, block_size, total_size):
        percentage = min(int(block_num * block_size * 100 / total_size), 100)
        self.item_progress.set(percentage)

    def close(self):
        match self.status:
            case 'Scanning':
                if (messagebox.askyesno('Exit', 'There is an active scan, Are you sure you want to quit?') == True):
                    self.status = 'Cancel'
            case 'Scanned':
                if (messagebox.askyesno('Exit', 'You will lose your current scan, Are you sure you want to quit?') == True):
                    self.status = 'Cancel'
            case 'Downloading':
                if (messagebox.askyesno('Exit', 'A download is in progress, Are you sure you want to quit?') == True):
                    self.status = 'Cancel'
            case _:
                self.status = 'Cancel'

        if (self.status == 'Cancel'):
            self.root.destroy()
            exit()

    def show(self, widget):
        widget.grid()

    def hide(self, widget):
        widget.grid_remove()

root = Tk()
window = AsanaExport(root)
root.mainloop()
