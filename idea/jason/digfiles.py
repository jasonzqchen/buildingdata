# -*- coding: utf-8 -*-
"""
Created on Sat Jul 28 14:05:23 2018

@author: jason.chen
@contact: jason.chen@arup.com
@company: ARUP

"""
import os 
import re 
import time 
import shutil

os.getcwd()


# %%   
def _subdirs(path):
    """Yield directory names not starting with '.' under given path."""
    try:
        for entry in os.scandir(path):
            if not entry.name.startswith('.') and entry.is_dir():
                yield entry
    except PermissionError as reason:
            pass #print('No permission'+ str(reason))
    except:
            pass #print('Error..')
            
def _subfiles(path):
    """Yield directory names not starting with '.' under given path."""
    try:
        for entry in os.scandir(path):
            if not entry.name.startswith('.') and entry.is_file():
                yield entry
    except PermissionError as reason:
            pass #print('No permission'+ str(reason))
    except:
            pass #print('Error..')
 
def _createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def _copyfile(store, savepath, overwrite = 0):
    try:
        for each in store:
            #print('each: ', each)
            directory = savepath + os.sep + each.split('\\')[-1]
            if not os.path.exists(directory):
                shutil.copy2(each, savepath)
            elif os.path.exists(directory):
                if overwrite == 0:
                    temp = '(1)' + each.split('\\')[-1]
                    modifiedpath = savepath + os.sep + temp
                    shutil.copy2(each, modifiedpath)
                else:
                    shutil.copy2(each, savepath) ## can this overwrite??
    except OSError:
        print ('Error: Creating directory. ' +  savepath)
        
def _loop(path, store):
    for eachEntry in _subdirs(path):
        if re.search(folderkey, eachEntry.name, re.IGNORECASE): #search the shallow layer
            _loop(eachEntry.path, store) # search the deeper layer
            #print('EPR: ', eachEntry.path)
            for eachFile in _subfiles(eachEntry.path):
                if not eachFile.name.startswith('~') and \
                re.search(filekey, eachFile.name, re.IGNORECASE): # this is to find the files, to be modified
                    store.append(eachFile.path)
    for eachFile in _subfiles(path):
            if not eachFile.name.startswith('~') and \
            re.search(filekey, eachFile.name, re.IGNORECASE): # this is to find the files, to be modified
                store.append(eachFile.path)   
    return store

def _loopfolders(paths):
    folders = []
    for path in paths:
        for eachEntry in _subdirs(path):
            if eachEntry.is_dir():
                folders.append(eachEntry.path)
    return folders

def _loopfoldersName(paths):
    names = []
    for path in paths:
        for eachEntry in _subdirs(path):
            if eachEntry.is_dir():
                names.append(eachEntry.name)
    return names

def _findfolders(folders):
    newpaths = []
    for each in folders:
        if re.search(folderkey, each, re.IGNORECASE):
            newpaths.append(each)
    return newpaths

def _looplooploop(newpaths):
    listoffolders = _loopfolders(newpaths)
    newpaths = _findfolders(listoffolders)
    #print(newpaths)
    return listoffolders, newpaths
        
def _slowfindfile(filekey, path):
    result = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if re.search(filekey,file,re.IGNORECASE) is not None:
                result.append(os.path.join(root, file))
    return result
        
def _main(path, full):
    store = []
    newpaths = []
    while True:
        if store != []: 
            # if there's folderkey found in the layer
            # this speeds up the searching logic a lot, however, may have chance to miss
            # the file if the wanted files are not put into the folder searched by
            # folderkey.
            store = _loop(path, store) # step 1
            break
        else:
            # if there no folderkey in the layer
            newpaths = []
            newpaths.append(path)
            counter = 0
            while True:
                # here is to loop to search the folder based on keywords
                (listoffolders, newpaths) = _looplooploop(newpaths)
                if newpaths == []:
                    (listoffolders, newpaths) = _looplooploop(listoffolders)
                else:
                    (listoffolders, newpaths) = _looplooploop(newpaths)
                    
                counter += 1
                if counter > 100:  #ok...if cannot search through finding folderkey...then let's do os.walk..
                                    # through all files....
                    if full:
                        store = _slowfindfile(filekey, path)
                    return store
                 
                # after we find the paths for folders..let's continue
                if len(newpaths) == 1:
                    store = _loop(newpaths[0], store)
                    return store
                if len(newpaths) > 1:
                    for newpath in newpaths:
                        #print(newpath)
                        store = _loop(newpath, store)
                    return store
                
def _input(path, name, datadict, savepath = [], full=0, save=0, overwrite=0):                
    #start = time.time()
    global folderkey, filekey
    folderkey = '(超限)|(抗震)|(EPR)|(报告)|(Report(s)?)'
    #folderkey = u'(Report(s)?)|(Calc)|(楼面)|(抗震)|(EPR)'
    filekey = '(超限)|(抗震)|(EPR)'
    #filekey = u'(楼面体系)|(钢筋桁架)|(组合楼板)'
    
    # 1. 用_loop找path，找到文件夹关键字（如folder）就继续搜索一直搜到有文件名关键字为止
    # 2. 如果step 1找不到return[]，要深度搜索直至搜到文件夹关键字为止，跳回step 1
    store = _main(path, full)
    #print('\n' + name)
    #print(store)
    if store == []:
        pass #print('Warning: cannot find the files, please review the targeted folder or filekey')
    else:
        datadict[name] = [store]
    
    #print('%.2f %s' % ((time.time()-start),'s')) 
    #start = time.time()
    if save and store !=[]:
        _createFolder(savepath + os.sep + name)
        savepathfull = savepath + os.sep + name
        print('copying...', name)
        _copyfile(store, savepathfull, overwrite)
    #print('%.2f %s' % ((time.time()-start),'s'))
# %%  
if __name__ == '__main__':
    start = time.time()
    
    paths = [r'G:',r'N:\STR-SHANTS02\项目',r'N:\STR-SHANTS03']
    #paths = [r'G:\251242 - CRC SHA SHW']
    savepath = r'C:\Users\jason.chen\Documents\JASON\EPR'
    
    datadict = {}
    projectpaths = _loopfolders(paths)
    projectnames = _loopfoldersName(paths)
    for projectpath, projectname in zip(projectpaths, projectnames):
        _input(projectpath, projectname, datadict, savepath, 1, 1, 0)

    print(time.time() - start)

# %% 从头到尾深度搜索
start = time.time()    

#filekey = u'(楼面体系)|(钢筋桁架)|(组合楼板)'
#paths = [r'G:', r'N:\STR-SHANTS02\项目', r'N:\STR-SHANTS03']
folderkey = '(超限)|(抗震)|(EPR)|(报告)|(Report(s)?)'
filekey = '(超限)|(抗震)|(EPR)'
datadict = {}
projectpaths = _loopfolders(paths)
projectnames = _loopfoldersName(paths)

for projectpath, projectname in zip(projectpaths, projectnames):
    result = _slowfindfile(filekey, projectpath)
    if result != []:
        datadict[projectname] = result
    
print(time.time() - start)
