import os, errno, urllib, subprocess, shutil, time
from win32com.client import Dispatch
import xml.etree.ElementTree as ET

MPpath = os.getenv("MaestroPanelPath")
WebcfgPath = "%s\Web\www\Web.config" % MPpath
corePath = r"%s\Web\www\bin\maestropanel.web.dll" % MPpath
MPAgentPath = "%s\Agent\MstrSvc.exe" % MPpath
MPappPath = "%s\Web\service\MstrW3Svc.exe" % MPpath
ver_parser = Dispatch('Scripting.FileSystemObject')

if os.path.isfile(MPappPath):
    MPAppFileVer = ver_parser.GetFileVersion(r'%s'% MPappPath)
if os.path.isfile(MPAgentPath):
    AgentFileVer = ver_parser.GetFileVersion(r'%s'% MPAgentPath)
if os.path.isfile(corePath):
    coreVer = ver_parser.GetFileVersion(r'%s'% corePath)
updatesDir = "%s\updates" % MPpath
repourl = "http://repo.maestropanel.com/A1/"

def saat():
    return time.strftime("%H:%M:%S: ")

def check_dir(path):
    try:
        if os.path.exists(path):
            shutil.rmtree(path)
        os.makedirs(path)
    except OSError, exc:
        if exc.errno != errno.EEXIST:
            raise
check_dir(updatesDir)

urllib.urlretrieve("http://repo.maestropanel.com/A1/pyupdates.xml", ("%s\pyupdates.xml" % updatesDir))

tree = ET.parse((r"%s\pyupdates.xml" % updatesDir))
root = tree.getroot()
verlist = {}
last_ver = 0
for updateData in root.findall('ver'):

    appver = updateData.find('appver').text
    name = updateData.get('name')
    link = updateData.find('link').text
    updateFile = updateData.find('updateFile').text
    if int(name) > last_ver:
        last_ver = int(name)
    #print name, appver, link, updateFile
    verlist[int(name)] = str(appver), str(updateFile)



def curr_vers():
    tempver = ""
    if os.path.isfile(corePath) == False:
        if os.path.isfile(MPAgentPath) == True:
            for i in verlist:
                if verlist[i][0] == AgentFileVer:
                    tempver = i
    if os.path.isfile(corePath) == True:
       for i in verlist:
           if verlist[i][0] == coreVer:
                tempver = i
    return tempver




def dl_update(x):
    updatever = os.path.join(updatesDir, x)
    dl_link = "%s%s" % (repourl, x)
    if not os.path.exists(updatever):
        print saat(), "Downloading %s" % x
        urllib.urlretrieve(dl_link, updatever)
    if os.path.exists(updatever):
        pass



def updateproc():

    currentversion = curr_vers()
    nextversion = currentversion + 1
    if currentversion == last_ver:
        print saat(),"Uygulama guncel."
    if currentversion != last_ver:
        print saat(), "Mevcut Versiyon %s" % verlist[curr_vers()][0]
        print saat(), "Uygulama son surum olan %s versiyonuna yukseltilecek" % verlist[last_ver][0]
        while currentversion != last_ver:
            #dl_update(verlist[nextversion][1])
            print saat(), "Guncelleme surumu %s yuklenecek." % verlist[nextversion][0]
            #subprocess.call("%s\%s /SILENT /SUPPRESSMSGBOXES" % (updatesDir, verlist[nextversion][1]))
            nextversion += 1
            currentversion += 1
            if currentversion == last_ver:
                print saat(), "Guncelleme Islemi bitti. Uygulama Surumu %s" % verlist[currentversion][0]

updateproc()
raw_input()