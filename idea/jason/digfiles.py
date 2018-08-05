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

# %% 
class Digfiles:
    def __init__(self, paths, savepath):    
        self.paths = paths # main search master path
        self.savepath = savepath # save to which path
        self.folderkey = [] # the keywords for searching prioritised folders
        self.filekey = [] # the keywords for searching prioritised files
        self.datadict = dict() # dictionary to store all the files for all projects
    
    def runsetting(self, save=0, overwrite=0, semiwalk= 0, fullwalk=0):
        '''
        ******run settings*******
        save: save the found files to the target folder
        overwrite: save the found files to the target folder and overwrite 
                    if there's same files
        semiwalk: semiwalk is in between keywords and full os.walk method, it use os.walk
                    when cannot use keywords to find the file
        fullwalk: call os.walk functions (slowfindfile) for full path search, very slow..
        
        '''
        self.save = save
        self.overwrite = overwrite
        self.semiwalk = semiwalk
        self.fullwalk = fullwalk
    
    def setfolderkey(self, key):
        self.folderkey = re.compile(key, re.IGNORECASE)

    def setfilekey(self, key):
        self.filekey = re.compile(key, re.IGNORECASE)
        
    def subdirs(self, path):
        """Yield directory names not starting with '.' under given path."""
        try:
            for entry in os.scandir(path):
                if not entry.name.startswith('.') and entry.is_dir():
                    yield entry
        except PermissionError as reason:
                pass #print('No permission'+ str(reason))
        except:
                pass #print('Error..')
                
    def subfiles(self, path):
        """Yield directory names not starting with '.' under given path."""
        try:
            for entry in os.scandir(path):
                if not entry.name.startswith('.') and entry.is_file():
                    yield entry
        except PermissionError as reason:
                pass #print('No permission'+ str(reason))
        except:
                pass #print('Error..')
#----------------------------------------------------------------------------      
    # this is a recursive function to update path and store all the files found            
    def loop(self, path, store):
        entries = Digfiles.subdirs(self, path)
        for eachEntry in entries:
            if self.folderkey.search(eachEntry.name): #search the shallow layer
                Digfiles.loop(self, eachEntry.path, store) # search the deeper layer
                
        files = Digfiles.subfiles(self, path)
        for eachFile in files:
                if not eachFile.name.startswith('~') and \
                self.filekey.search(eachFile.name): # this is to find the files, to be modified
                    store.append(eachFile.path)   
        return store
    
#----------------------------------------------------------------------------  
    def loopfoldersName(self, paths):
        names = []
        for path in paths:
            entries = Digfiles.subdirs(self, path)
            for eachEntry in entries:
                names.append(eachEntry.name)
        return names
    
    def loopfolders(self, paths):
        for path in paths:
            folders = [eachEntry.path for eachEntry in Digfiles.subdirs(self, path)]
            files = [eachFile.path for eachFile in Digfiles.subfiles(self, path)]
        return folders, files

    def findfolders(self, folders, files):       
        ffolders = [folder for folder in folders if self.folderkey.search(folder)]
        ffiles = [file for file in files if self.filekey.search(file)]
        return ffolders, ffiles
    
    def ziploop(self, newpaths):
        (listoffolders, listoffiles) = Digfiles.loopfolders(self, newpaths)
        (newpaths, ffiles) = Digfiles.findfolders(self, listoffolders, listoffiles)
        #print(newpaths)
        return listoffolders, listoffiles, newpaths, ffiles
#----------------------------------------------------------------------------
    # others 
    def createFolder(self, directory):
        '''
        this function can be used to create folders
        '''
        try:
            if not os.path.exists(directory):
                os.makedirs(directory)
        except OSError:
            print ('Error: Creating directory. ' +  directory)
    
    def copyfile(self, files, saveto):
        '''
        this function can be used to copy files
        files: a list of file paths
        saveto: a path that all files will be copied to
        '''
        try:
            for each in files:
                #print('each: ', each)
                fullpath = saveto + os.sep + each.split('\\')[-1]
                if not os.path.exists(fullpath):
                    shutil.copy2(each, saveto)
                elif os.path.exists(fullpath):
                    if self.overwrite == 0:
                        temp = '(1)' + each.split('\\')[-1]
                        modifiedpath = saveto + os.sep + temp
                        shutil.copy2(each, modifiedpath)
                    else:
                        shutil.copy2(each, saveto) ## can this overwrite??
        except OSError:
            print ('Error: Creating directory. ' +  savepath)     
            
    def slowfindfile(self, path):
        '''
        this function use os.walk to walk through all directories in path, and
        return with a stored function storing all the file path that meet the
        filekey search
        
        use with cautions -> very slow
        '''
        store = []
        for root, dirs, files in os.walk(path):
            for file in files:
                if self.filekey.search(file) is not None:
                    store.append(os.path.join(root, file))
        return store
    
#----------------------------main execution code---------------------------------

    def main(self, path):
        '''
        main function to define logic to find the targeted folders and files
        Basic logic:
        Three senerios:
        1. if there's folderkey found in the layer.
            this speeds up the searching logic a lot; however, may have chance to miss
            the file if the wanted files are not put into the folder searched by
            folderkey.
        2. if there no folderkey in the layer, then go to deeper layer to find
        3. semi-walking is in between keywords and full os.walk method, it use os.walk
            when cannot use keywords to find the file
        4. call os.walk functions (slowfindfile) for full path search, very slow..
                
        '''
        store = []
        
        while True:
            #print('method 1...shallow.digging...', path)
            store = Digfiles.loop(self, path, store) # step 1
            if self.fullwalk:
                #print('method 4...fullwalking....', path)
                store = Digfiles.slowfindfile(self, path)
                return store
            elif store != []: 
                # if there's folderkey found in the layer
                # this speeds up the searching logic a lot, however, may have chance to miss
                # the file if the wanted files are not put into the folder searched by
                # folderkey.
                return store
            else:
                #print('method 2...deep.digging..', path)
                # if there no folderkey in the layer, go deeper
                newpaths = []
                newpaths.append(path)
                counter = 0
                
                (listoffolders, lfis, newpaths, ffis) = Digfiles.ziploop(self, newpaths)
                
                while True:
                    # here is to loop to search the folder based on keywords
                    if newpaths == []: 
                        # if the first search didn't get the anything, then we need to
                        # feed back the all the folderpath back to loop and find the
                        # next layer of folders
                        if listoffolders != []:
                            (listoffolders, lfis, newpaths, ffis) = Digfiles.ziploop(self, listoffolders)
                            if ffis != []:
                                for ffi in ffis:
                                    store.append(ffi)
                    if len(newpaths) == 1:
                        store = Digfiles.loop(self, newpaths[0], store)
                        return store
                    if len(newpaths) > 1:
                        for newpath in newpaths:
                            #print(newpath)
                            store = Digfiles.loop(self, newpath, store)
                        return store
                    
                    counter += 1
                    if counter > 100:  #ok...if cannot search through finding folderkey...then let's do os.walk..
                                        # through all files....
                        if self.semiwalk:
                            #print('method 3...deep.digging..', path)
                            store = Digfiles.slowfindfile(self, path)
                            #print('\n',path,'\n')
                            #print(store)
                        return store
                
    def store2dict(self, path, name):             
        '''
        store the data into dictionary with each main folder name
        and copy to the target folder
        '''
        store = Digfiles.main(self, path)
        #print('\n' + name)
        #print(store)
        if store == []:
            pass #print('!Warning: cannot find the files, please review the targeted folder or filekey')
        else:
            self.datadict[name] = [store]
        
        if self.save and store !=[]:
            Digfiles.createFolder(self, self.savepath + os.sep + name)
            savepathfull = self.savepath + os.sep + name
            print('copying...', name)
            Digfiles.copyfile(self, store, savepathfull)
    
    def run(self):
        (projectpaths,files) = df.loopfolders(self.paths)
        projectnames = df.loopfoldersName(self.paths)
        for projectpath, projectname in zip(projectpaths, projectnames):
            Digfiles.store2dict(self, projectpath, projectname)

        
if __name__ == '__main__':        
    start = time.time()
    
    paths = paths = [r'C:\Users\Ziqiang\OneDrive - Arup']
    savepath = r'C:\Users\Ziqiang\Desktop\EPR'
    df = Digfiles(paths, savepath)
    
    folderkeyword = '(超限)|(抗震)|(EPR)|(报告)|(Report(s)?)'
    filekeyword = '(超限)|(抗震)|(EPR)'
    df.setfolderkey(folderkeyword)
    df.setfilekey(filekeyword)    
    df.runsetting(0,0,1,0)
    # save
    # overwrite
    # semiwalk
    # fullwalk
    
    df.run()
    filedict = df.datadict
    
    print(time.time() - start)