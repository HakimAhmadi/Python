import win32file
import win32ui 
import win32con
import win32gui 
import win32api
from ctypes import windll
from PIL import Image
import time


def fullWindowScreenshot():
    # grab a handle to the main desktop window 
    hdesktop = win32gui.GetDesktopWindow()
    
    # determine the size of all monitors in pixels 

    width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN) 
    height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN) 
    left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN) 
    top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN) 

    # create a device context 
    desktop_dc = win32gui.GetWindowDC(hdesktop) 
    img_dc = win32ui.CreateDCFromHandle(desktop_dc) 
    
    # create a memory based device context 
    mem_dc = img_dc.CreateCompatibleDC() 
    
    # create a bitmap object 
    screenshot = win32ui.CreateBitmap() 
    screenshot.CreateCompatibleBitmap(img_dc, width, height) 
    mem_dc.SelectObject(screenshot) 
    
    # copy the screen into our memory device context 
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top),win32con.SRCCOPY) 
    
    # save the bitmap to a file 
    screenshot.SaveBitmapFile(mem_dc, '...screenshot.bmp') 
    # free our objects 
    mem_dc.DeleteDC() 
    win32gui.DeleteObject(screenshot.GetHandle()) 


def singleAppScreenshot():
        
    hwnd = win32gui.FindWindow(None, 'Calculator')
    
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bot - top

    print(left, top, right, bot)
    hwndDC = win32gui.GetWindowDC(hwnd)
    # print("hwnDC",hwndDC)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    # print("mfcDC",mfcDC)
    saveDC = mfcDC.CreateCompatibleDC()
    # print("saveDC",saveDC)

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    # result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        im.save("test.png")

    
    # # hwnd_target = 0x00480362 #Chrome handle be used for test 
    # hwnd_target = win32gui.FindWindow(None, 'Calculator')

    # left, top, right, bot = win32gui.GetWindowRect(hwnd_target)
    # w = right - left
    # h = bot - top

    # win32gui.SetForegroundWindow(hwnd_target)
    # time.sleep(1.0)

    # hdesktop = win32gui.GetDesktopWindow()
    # hwndDC = win32gui.GetWindowDC(hdesktop)
    # mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    # saveDC = mfcDC.CreateCompatibleDC()

    # saveBitMap = win32ui.CreateBitmap()
    # saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    # saveDC.SelectObject(saveBitMap)

    # result = saveDC.BitBlt((0, 0), (w, h), mfcDC, (left, top), win32con.SRCCOPY)

    # bmpinfo = saveBitMap.GetInfo()
    # bmpstr = saveBitMap.GetBitmapBits(True)

    # im = Image.frombuffer(
    #     'RGB',
    #     (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
    #     bmpstr, 'raw', 'BGRX', 0, 1)

    # win32gui.DeleteObject(saveBitMap.GetHandle())
    # saveDC.DeleteDC()
    # mfcDC.DeleteDC()
    # win32gui.ReleaseDC(hdesktop, hwndDC)

    # if result == None:
    #     #PrintWindow Succeeded
    #     im.save("test.png")


