import os, datetime, time, webbrowser, win32api, win32con, getpass, win32com.client

shell = win32com.client.Dispatch("WScript.Shell")
password = "Vizard321"
key = getpass.getpass('DNS Error...')
op = getpass.getpass('Op?')
cx = 1516
cy = 131
yx = 910
yy = 597
nx = 1005
ny = 597
mx = 1669
my = 125
sx = 1662
sy = 202
sx2 = 1667
sy2 = 249
ax = ''
ay = ''


if key == password and op == "o":
    do = int(input("D: "))
    ho = int(input("H: "))
    mo = int(input("M: "))
    so = 00

    out_time = datetime.datetime(2019, 1, do, ho, mo, so)

    while(True):
        dtn = datetime.datetime.now()
        if dtn >= out_time:
            os.system("TASKKILL /F /IM vizdbupdate.exe")
            time.sleep(2)
            shell.Run('Chrome')
            time.sleep(2)
            os.system("TASKKILL /F /IM chrome.exe")
            time.sleep(2)
            shell.Run('Chrome')
            time.sleep(2)
            shell.SendKeys("% ", 0)
            time.sleep(2)
            shell.SendKeys("x", 0)
            time.sleep(2)
            shell.SendKeys("http://bolportal.myboltv.net:200/EmployeePortalV3/", 0)
            time.sleep(10)
            win32api.SetCursorPos((cx,cy))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,cx,cy,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,cx,cy,0,0)
            time.sleep(2)
            win32api.SetCursorPos((yx,yy))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,yx,yy,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,yx,yy,0,0)
            time.sleep(5)
            win32api.SetCursorPos((nx,ny))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,nx,ny,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,nx,ny,0,0)
            time.sleep(5)
            win32api.SetCursorPos((mx,my))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,mx,my,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,mx,my,0,0)
            time.sleep(2)
            win32api.SetCursorPos((sx,sy))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,sx,sy,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,sx,sy,0,0)
            time.sleep(2)
            win32api.SetCursorPos((yx,yy))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,yx,yy,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,yx,yy,0,0)
            time.sleep(5)
            os.system("TASKKILL /F /IM chrome.exe")
            break
        else:
            break

if key == password and op == "i":
    un = getpass.getpass('Who?')
    pp = getpass.getpass('How?')
    di = int(input("D: "))
    hi = int(input("H: "))
    mi = int(input("M: "))
    si = int(input("S: "))

    end_time = datetime.datetime(2019, 1, di, hi, mi, si)

    while(True):
        dtn = datetime.datetime.now()
        if dtn >= end_time:
            #os.system("TASKKILL /F /IM chrome.exe")
            #time.sleep(5)
            shell.Run('Chrome')
            time.sleep(2)
            shell.SendKeys("% ", 0)
            time.sleep(5)
            shell.SendKeys("x", 0)
            time.sleep(2)
            shell.SendKeys("http://bolportal.myboltv.net:200/EmployeePortalV3/", 0)
            time.sleep(5)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(2)
            shell.SendKeys("^a", 0)
            time.sleep(2)
            shell.SendKeys("{DEL}", 0)
            time.sleep(2)
            shell.SendKeys(un, 0)
            time.sleep(2)
            shell.SendKeys("{TAB}", 0)
            time.sleep(2)
            shell.SendKeys("{DEL}", 0)
            time.sleep(2)
            shell.SendKeys(pp, 0)
            time.sleep(2)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(4)
            shell.SendKeys("{BKSP}", 0)
            time.sleep(2)
            shell.SendKeys("{F5}", 0)
            time.sleep(2)
            win32api.SetCursorPos((mx,my))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,mx,my,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,mx,my,0,0)
            time.sleep(2)
            win32api.SetCursorPos((sx2,sy2))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,sx2,sy2,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,sx2,sy2,0,0)
            time.sleep(2)
            win32api.SetCursorPos((yx,yy))
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,yx,yy,0,0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,yx,yy,0,0)
            break
