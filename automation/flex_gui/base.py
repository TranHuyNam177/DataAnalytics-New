import os
import re
import pywintypes
import pywinauto.timings
from pywinauto.application import Application
import cv2 as cv
import numpy as np


class Flex:

    def __init__(self):

        self.app = Application()
        self.mainWindow = None
        self.loginWindow = None
        self.funcCode = None
        self.funcWindow = None
        self.username = None
        self.password = None

    def start(self,existing:bool,**kwargs):
        os.chdir(r'../../FLEX_PROD') # cd vào FLEX_PROD
        if existing:
            self.app = self.app.connect(title_re='^\.::.*Flex.*',**kwargs)
        else:
            self.app = self.app.start(cmd_line='@DIRECT.exe')
        self.mainWindow = self.app.window(title_re='.*Flex.*')
        self.loginWindow = self.app.window(title='Login')
        if os.path.isdir(r'../dist/VCI1104'):
            os.chdir(r'../dist/VCI1104')
        else:
            os.chdir(r'../AutomationApp/VCI1104')

    def login( # tested
        self,
        username,
        password,
    ):
        # Override
        setFocus(self.mainWindow)
        self.username = username
        self.password = password
        # Nhập User name
        usernameBox = self.loginWindow.child_window(auto_id='txtUserName')
        usernameBox.click()
        usernameBox.type_keys('^a{DELETE}')
        usernameBox.type_keys(username)
        # Nhập Password
        passwordBox = self.loginWindow.child_window(auto_id='txtPassword')
        passwordBox.click()
        passwordBox.type_keys('^a{DELETE}')
        passwordBox.type_keys(password+'{ENTER}'*2)
        # Cập nhật phiên bản (nếu có)
        updateBox = self.app.window(title='Update')
        if updateBox.exists(timeout=5,retry_interval=0.5):
            print('Update new version...')
            updateBox.type_keys('{ENTER}')
            self.mainWindow.wait_not('visible',timeout=30,retry_interval=0.5) # đợi trình cập nhật tự tắt Flex
            # connect vào Flex (đã được tự mở lên sẵn)
            try:
                self.start(existing=True,timeout=30) # đợi tối đa 30s cho Flex cập nhật xong
                print('Connecting to existing Flex...')
            except (pywinauto.timings.TimeoutError,): # đợi 30s mà ko tự mở lên được (do lỗi Flex)
                self.start(existing=False) # start Flex mới
                print('Failed to connect to existing Flex, starting new Flex instance...')
            self.login(username,password) # Đăng nhập lại
        self.loginWindow.wait_not('exists',timeout=5,retry_interval=0.5)

    def getFuncCode(self):
        return self.funcCode

    def setFuncCode(self,funcCode):
        self.funcCode = funcCode

    def insertFuncCode(self,funcCode):
        setFocus(self.mainWindow)
        funcBox = self.mainWindow.children()[-3]
        funcColorImg = np.array(funcBox.capture_as_image())
        funcGrayImg = cv.cvtColor(funcColorImg,cv.COLOR_RGB2GRAY)
        funcGrayImg[:,:50] = 0 # đè màu đen để che logo Flex đi, tránh bắt nhầm màu trắng trong logo
        _, funcBinaryImg = cv.threshold(funcGrayImg,254,255,cv.THRESH_BINARY)
        locTuple = np.nonzero(funcBinaryImg==255)
        yLoc = np.median(locTuple[0]).astype(np.uint8)
        xLoc = np.median(locTuple[1]).astype(np.uint8)
        funcBox.click_input(coords=(xLoc,yLoc))
        funcBox.type_keys('^a{DELETE}')
        funcBox.type_keys(funcCode+'{ENTER}')
        self.funcWindow = self.app.window(title_re=f".*{re.sub('[A-Z]*','',funcCode)}.*")
        self.funcWindow.maximize()
        self.setFuncCode(funcCode)

    def exitAllWindows(self):
        appWindows = self.app.windows()
        for appWindow in appWindows:
            if appWindow.is_visible() and appWindow != self.mainWindow:
                appWindow.close()
        setFocus(self.mainWindow)


def setFocus(window):
    window.maximize()
    while True:
        try:
            window.set_focus()
            break
        except (pywintypes.error,):
            print('Click again')
