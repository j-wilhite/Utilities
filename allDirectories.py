__author__ = 'Jessica'

import os, sys, getopt
import subprocess
import win32com.client.gencache

def main(argv):

    try:
      opts, args = getopt.getopt(argv[:],'hc:',['help', 'command='])
    except getopt.GetoptError:
      print('allDirectories.py -c <command> /nsave: saves paths of all open directories/nopen: opens saved directories in windows explorer')
      sys.exit(2)
    for opt, arg in opts:
        if opt in ("-c", "--command"):
            commandOption = arg
            if commandOption == "save":
                saveOpenDirectories()
            elif commandOption == "open":
                directoryFilePath = getDirectoriesFile()
                if directoryFilePath:
                    import pdb;pdb.set_trace()
                    openSavedDirectories(directoryFilePath)

def getDirectoriesFile():
    dirText = os.getcwd()+ "\\AllOpenDirectories.txt"
    return dirText

def saveOpenDirectories():
    allOpenDirectories = []
    for w in win32com.client.gencache.EnsureDispatch("Shell.Application").Windows():
        allOpenDirectories.append(w.LocationURL.strip('file:///')+"\n")

    dirText = open(os.getcwd()+ "\\AllOpenDirectories.txt", "w")
    dirText.writelines(allOpenDirectories)
    dirText.close()

def openSavedDirectories(dirText):
    if os.path.exists(dirText):
        savedDirs = open(dirText).readlines()
        for directoryPath in savedDirs:

            subprocess.Popen('explorer ' + directoryPath.strip('\n').replace('/', '\\'))


if __name__ == "__main__":
   main(sys.argv[1:])