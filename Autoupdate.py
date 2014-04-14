import os, errno, urllib, subprocess
from win32com.client import Dispatch

verlist = {     1   : ["0.9.2.19", "Update093.exe"],
                2   : ["0.9.3.0", "Update093.exe"],
                3   : ["0.9.3.25", "Update093_25.exe"],
                4   : ["0.9.3.27", "Update093_27.exe"],
                5   : ["0.9.3.32", "Update093_32.exe"],
                6   : ["0.9.3.33", "Update093_33.exe"],
                7   : ["0.9.3.92", "Update093_92.exe"],
                8   : ["0.9.3.126", "Update093_126.exe"],
                9   : ["0.9.3.132", "Update093_132.exe"],
                10  : ["0.9.3.147", "Update093_147.exe"],
                11  : ["0.9.4.0", "Update094.exe"],
                12  : ["0.9.4.15", "Update094_15.exe"],
                13  : ["0.9.4.218", "Update094_218.exe"]
}
MPpath = os.getenv("MaestroPanelPath")
WebcfgPath = "%s\Web\www\Web.config" % MPpath
corePath = r"%s\Web\www\bin\maestropanel.core.dll" % MPpath
MPAgentPath = "%s\Agent\MstrSvc.exe" % MPpath
MPappPath = "%s\Web\service\MstrW3Svc.exe" % MPpath
ver_parser = Dispatch('Scripting.FileSystemObject')
AgentFileVer = ver_parser.GetFileVersion(r'%s'% MPAgentPath)
MPAppFileVer = ver_parser.GetFileVersion(r'%s'% MPappPath)
coreVer = ver_parser.GetFileVersion(r'%s'% corePath)
updatesDir = "%s\updates" % MPpath
repourl = "http://repo.maestropanel.com/A1/"

def check_dir(path):
    try:
        os.makedirs(path)
    except OSError, exc:
        if exc.errno != errno.EEXIST:
            raise
check_dir(updatesDir)

def curr_vers():
    tempver = ""
    if os.path.isfile(corePath) == False:
        print "MaestroPanel uygulamasi bulunamadi."
    if os.path.isfile(corePath) == True:
       for i in verlist:
           if verlist[i][0] == coreVer:
                tempver = i
    return tempver

def dl_update(x):
    updatever = os.path.join(updatesDir, x)
    dl_link = "%s%s" % (repourl, x)
    if not os.path.exists(updatever):
        print "Downloading %s" % x
        urllib.urlretrieve(dl_link, updatever)
    if os.path.exists(updatever):
        pass

print "Mevcut Versiyon %s" % verlist[curr_vers()][0]

def updateproc():
    finalversion = curr_vers() + 1
    if finalversion == 14:
        print "uygulama son surum"
    if finalversion < 14:
        while finalversion < 14:
            dl_update(verlist[finalversion][1])
            print "Will run %s\%s /SILENT /SUPPRESSMSGBOXES" % (updatesDir, verlist[finalversion][1])
            subprocess.call("%s\%s /SILENT /SUPPRESSMSGBOXES" % (updatesDir, verlist[finalversion][1]))
            finalversion += 1

updateproc()

#print verlist[4][1]
#os.path.isfile(MPappPath) == True:
#subprocess.call("C:\updates\update94.exe /SILENT /SUPPRESSMSGBOXES")

#eski curr_vers
"""    tempver = ""
    tempver2 = ""
    if os.path.isfile(MPAgentPath) == True:
       for i in verlist:
           if verlist[i][0] == AgentFileVer:
                tempver = i
    if os.path.isfile(MPappPath) == True:
       for i in verlist:
           if verlist[i][0] == MPAppFileVer:
                tempver2 = i
    if tempver > tempver2:
        return tempver2
    else:
        return tempver"""