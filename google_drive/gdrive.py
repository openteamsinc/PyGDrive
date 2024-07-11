import os
import pathlib

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from apiclient.http import MediaFileUpload

from .drive_utils import DriveFileNotFoundError, DuplicateFileName

current_path = resourcing_path=(pathlib.Path(__file__).parent.resolve())

class GoogleDriveInstance():
    def __init__(self, ini_path=current_path, api='drive', api_version='v3'):
        self.ini_path = ini_path
        self.service = self.authenticate(ini_path, api, api_version)
        if api == 'drive':
            self.file_dict = self.list_files()
            self.folders = self.list_files(only_folders=True)
            self.files = self.list_files(only_files=True)
        
    def authenticate(self, ini_path, api, api_version):
        scope = ["https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(ini_path.joinpath("client_secrets.json"), scope)
        service = build(api, api_version, credentials=creds)
        return service
    
    def list_files(self, folder=None, only_files=False, only_folders=False):
        files = []
        query_params = []
        page_token = ""
        while True:
            if folder:
                folder_id_q = f"'{folder}' in parents"
                query_params.append(folder_id_q)
            
            mime_type_q = None
            if only_files:
                mime_type_q = "mimeType != 'application/vnd.google-apps.folder'"
            if only_folders:
                mime_type_q = "mimeType = 'application/vnd.google-apps.folder'"
            if mime_type_q:
                query_params.append(mime_type_q)
            
            query = " and ".join(query_params)
            
            response = self.service.files().list(
                        q=query,
                        pageToken=page_token,
                        includeItemsFromAllDrives=True,
                        supportsAllDrives=True,
                        fields="nextPageToken, files(id, name, parents, createdTime, mimeType)"
                    ).execute()
            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken')
            if not page_token:
                break
        
        file_dict = {}
        for file in files:
            if 'parents' not in file:
                file['parents'] = []

            file_dict[file['name']] = {
                'id': file['id'], 'parents': file['parents'],
                'created_time': file['createdTime'],
                'mime_type': file['mimeType'],
                'name':file['name']
            }

        return file_dict
    
    def update_file_list(self):
            
        self.files = self.list_files(only_files=True)
        self.folders = self.list_files(only_folders=True)  
    
    def upload_file(self, file_to_upload, parent_folder=None, return_data=False):
        # Uploads a file to specified google drive
        # Needs path to file or for file to be in current directory
        parent_id = self.get_file_id(parent_folder)
        file_metadata = {
            'name':file_to_upload.rsplit('/')[-1],
            'parents': [parent_id]
        }
        
        media = MediaFileUpload(os.path.expanduser(file_to_upload))

        try:
            return_data = self.service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        except FileNotFoundError:
            print('{} not found, please pass in either file in current dir or path to file'.format(file_to_upload))
        
        self.update_file_list()
        if return_data:
            return return_data

    def get_file_id(self, file_to_check, parent_id=None):
    # Takes a file name and checks if it is in the specified google drive
    # Returns a list of file IDs matching the provided file name
    # If parent_id is provided, only returns file IDs with that parent folder

        file_to_check = file_to_check.lower()
        matching_files = []

        for file_name, file_details in self.file_dict.items():
            if file_name.lower() == file_to_check:
                    if parent_id is None or parent_id in file_details["parents"]:
                        matching_files.append(file_details['id'])

        if len(matching_files) == 1:
            return matching_files[0]
        
        if matching_files is None:
            raise DriveFileNotFoundError(file_to_check)
        
        return matching_files


    
    def share_file(self, file_to_share, user_email, share_type='writer', share_all=False):
        # Shares a file by name or id with a user by email, default permission is writer

        file_id = self.get_file_id(file_to_share,get_all=share_all)
        user_permission = {
            'type':'user',
            'role':share_type,
            'emailAddress':user_email,
            'sendNotificationEmail':True
        }

        for i in file_id.split(', '):
            self.service.permissions().create(fileId=i, body=user_permission).execute()
            print("{} shared with {}".format(file_to_share, user_email))

    def remove_file(self, file_to_remove, remove_all=False):
        # Takes a file name or id and removes it from google drive
        # if remove all = True removes all files of that name
        
        file_id = self.get_file_id(file_to_remove, get_all=remove_all)
        
        for i in file_id.split(", "):
            self.service.files().delete(fileId=i,supportsAllDrives=True).execute()
            print(i + " deleted")
        
        self.update_file_list()

    def most_recent_file(self, parent_folder_id):
        file_dict = self.list_files(folder=parent_folder_id)
        most_recent_file = max(file_dict.values(), key=lambda x: x['created_time'])
        return most_recent_file