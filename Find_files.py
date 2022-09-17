import glob
import os
import pathlib
import pandas as pd
import shutil
folders="3MBeeline_OriginalFilePath,ADAC_OriginalFilePath,ADP_OriginalFilePath,Windstream_OriginalFilePath,WOODBRIDGE_OriginalFilePath,ZCHAOS_OriginalFilePath"
folders=folders.split(',')


def find_files(sourceid):
    list_of_files=[]
    
    for p in folders:
        path="//rp-tne-bt13file/"+p+"/"
        filepaths=[]
        for filepath in pathlib.Path(path).glob('**/*'):
            if sourceid in str(filepath):
                filepaths.append(str(filepath.absolute()))
        if len(list_of_files)==0:
            list_of_files = sorted( filepaths,key = os.path.getmtime)
    if len(list_of_files)!=0:
        print('copied:',sourceid)
        shutil.copy(list_of_files[-2:-1][0],r'K:/Users/vkharatmal/Downloads/sample files/Biztalk upgrade testing/TestFiles')
        return
    return(sourceid)
    
    
df=pd.read_excel('K:/Users/vkharatmal/Downloads/sample files/Biztalk upgrade testing/codes and settings/sourceids.xlsx')
df=df.values.tolist()
no_files=[]
for i in df:
    nofile=find_files(i[0])
    if nofile:
        no_files.append(nofile)
print('No files for:',no_files)
