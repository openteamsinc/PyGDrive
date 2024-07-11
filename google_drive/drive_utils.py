import gspread

class DuplicateFileName(Exception):
    # Error to display when more then one file of the same name are in the same google drive
    # Shows the IDs of all shared name files

    def __init__(self, failed_file, file_id):
        self.failed_file = failed_file
        self.file_id = file_id
        self.message = f"More then 1 file exists with name {failed_file}, please use a file id instead {file_id}"
        super().__init__(self.message)

    def __str__(self):
        return f'{self.message}'

class DriveFileNotFoundError(Exception):
    def __init__(self, not_found, parent_id=None):
        self.parent_id = parent_id
        self.not_found = not_found
        not_found_stmt = f'File with name/id "{not_found}"'
        parent_stmt = ''
        if parent_id:
            parent_stmt = f'and parent folder of id of "{parent_id}" '
        self.message = f'{not_found_stmt} {parent_stmt}not found in specified drive'
        super().__init__(self.message)

def get_column_by_header(ws, to_find, numeric=False):
    if numeric:
        return ws.find(to_find).index
    else:
        return ws.find(to_find).address[:-1]

def update_gsheet_columns(ws, form_dict, sh_len):
    for formula in form_dict:
        col_alph = get_column_by_header(ws, formula)
        formula_str = form_dict[formula]
        update_column_single_val(col_alph, sh_len, formula_str, ws)

def get_gsheet_by_name(sheet_name, file_path='client_secrets.json'):
    gc = gspread.service_account(file_path)
    sh = gc.open(sheet_name)
    return sh

def update_column_single_val(col_alpha, sh_len, value, worksheet):
    rows = sh_len
    values = [value]
    data = {'range':f'{col_alpha}2:{col_alpha}{sh_len+1}',
            'values':[values for row in range(rows)]}
    
    worksheet.batch_update([data], value_input_option = 'USER_ENTERED')
    print(f"finished update {data['range']}")

def yqm_folders(drive_instance, folder_name, today, division=None):
    division_name = f'{division}'
    month = today.month
    month_name = f'{today:%m}_{today:%Y}'
    quarter_name = f'Q{((month-1)//3)+1}'
    year_name = f"FY {today.year}"

    root_folder = drive_instance.get_file_id(folder_name)

    if division:
        division_folder = find_or_create_folder(drive_instance, division_name, root_folder)
        year_folder = find_or_create_folder(drive_instance, year_name, division_folder)
    else:
        year_folder = find_or_create_folder(drive_instance, year_name, root_folder)
        
    quarter_folder = find_or_create_folder(drive_instance, quarter_name, year_folder)
    month_folder = find_or_create_folder(drive_instance, month_name, quarter_folder)
    print(f'found folder {month_name}')
    return month_folder

def find_or_create_folder(drive_instance, folder_name, parent_folder):
    try:
        folder_id = drive_instance.get_file_id(folder_name, parent_id=parent_folder)
    except DriveFileNotFoundError:
        print(f'creating new folder {folder_name}')
        folder_metadata = {
            'name':folder_name,
            'parents':[parent_folder],
            'mimeType': 'application/vnd.google-apps.folder' 
        }
        file = drive_instance.service.files().create(body=folder_metadata,
                                                    fields='id',supportsAllDrives=True).execute()
        folder_id = file['id']
    return folder_id

def find_folder_id(drive_instance, folder_name, parent_folder):
    try:
        folder_id = drive_instance.get_file_id(folder_name, parent_id=parent_folder)
    except DriveFileNotFoundError:
        print(f'creating new folder {folder_name}')
        folder_metadata = {
            'name':folder_name,
            'parents':[parent_folder],
            'mimeType': 'application/vnd.google-apps.folder' 
        }
        file = drive_instance.service.files().create(body=folder_metadata,
                                                    fields='id',supportsAllDrives=True).execute()
        folder_id = file['id']
    return folder_id

def create_worksheets(sh, ws_list, remove_initial=True):
    ws_titles = [ws.title for ws in sh.worksheets()]

    for ws_name in ws_list:
        #ws_titles = [ws._properties['title'] for ws in sh.worksheets()]  
        if ws_name not in ws_titles:
            sh.add_worksheet(ws_name, 0, 0)
    
    if remove_initial and 'Sheet1' in ws_titles:
        sh.del_worksheet_by_id(0)