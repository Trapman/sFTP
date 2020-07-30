import os
import pysftp

host = 'abc.dev.google.com'
port = 22
username = 'mega.fund'
pw = 'AbCdEf123'

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None
output_dir = r"C:\Users\DAS\OneDrive - Tomo\Documents"     # this is just a OneDrive example
excel_array = []

with pysftp.Connection(host, username=username, password=pw, cnopts=cnopts) as sftp: 
    sftp.cwd('/outbox')      # this will most likely change depending on how the sFTP 'outbox' is named
    directory_structure = sftp.listdir_attr()
    
    for attr in directory_structure:
        
        extension = str(os.path.splitext(attr.filename)[1]).lower()
        print (attr.filename, attr)
        
        if extension == '.xlsx':
            excel_array.append(attr.filename)
            sftp.get(f'/outbox/{attr.filename}', f'{output_dir}\{attr.filename}')
