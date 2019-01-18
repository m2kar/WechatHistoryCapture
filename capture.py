"""
微信电脑版的聊天记录窗口打开，翻到最上面，然后程序会进行截图
"""
#   @Time:  2019/1/17 1:23
#   @File:  capture.py
#   @Author: Zhiqing Rui
import time
import os
import win32gui,win32ui,win32con,win32api
import mss
import mss.tools
import datetime

class WindowHide(Exception):
  pass

def get_hwnd():
  wclass=50103
  filename="1.bmp"
  hwnd=win32gui.FindWindow(wclass,"微信")
  return hwnd

def cap(filename,hwnd,head=180):
  # FileManagerWnd 50103
  l,t,r,b=win32gui.GetWindowRect(hwnd)
  if b<0 :
    raise WindowHide("wechat window ")
  ti=t+head
  # print(l,t,r,b)
  with mss.mss() as sct:
    monitor={"top":ti,"left":l,"width":r-l,"height":b-ti}
    sct_img=sct.grab(monitor)
    # output = ""
    mss.tools.to_png(sct_img.rgb,sct_img.size,output=filename)

def press_key(hwnd):
  win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NEXT, 0)
  pagedown=34
  win32api.keybd_event(pagedown,0,0,0)
  win32api.keybd_event(13,0,win32con.KEYEVENTF_KEYUP,0)

def cap_all():
  hwnd=get_hwnd()
  i=0
  curtime=datetime.datetime.now().strftime("%Y%m%d%H%M%S")
  dirname="capture"+curtime
  os.mkdir(dirname)
  # if os.path.exists(dir):
  #   if not os.path.isdir(dir):
  #     raise FileExistsError("该文件夹重名 {dir}".format(**locals()))
  # else:
  #   os.mkdir(dir)
  while True:
    try:
      # time.sleep()
      cap(os.path.join(dirname,"{i:05d}.png".format(i=i)),hwnd)
    except WindowHide:
      continue
    else:
      i+=1
      # time.sleep(0.1)
      win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_NEXT, 0)

if __name__ == '__main__':
    cap_all()
