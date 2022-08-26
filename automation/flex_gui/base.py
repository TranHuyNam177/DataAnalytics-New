import re
import os
import sys
import time
import pywintypes
from pywinauto.application import Application
import cv2 as cv
import numpy as np


class Flex:

    def __init__(self,flexEnv):

        if flexEnv.lower() == 'pro':
            self.flexExe = os.path.join(os.path.dirname(__file__),'Flex','FlexPro','@DIRECT.exe')
        elif flexEnv.lower() == 'uat7':
            self.flexExe = os.path.join(os.path.dirname(__file__),'Flex','FlexUAT7','@DIRECT.exe')
        else:
            raise ValueError('Invalid Flex Environtment')

        self.app = Application()
        self.mainWindow = None
        self.loginWindow = None
        self.funcCode = None
        self.funcWindow = None
        self.username = None
        self.password = None

    def start(self):
        self.app = self.app.start(self.flexExe)
        self.mainWindow = self.app.window(title_re='.*Flex.*')
        self.loginWindow = self.app.window(title='Login')

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
        time.sleep(5)

    def getFuncCode(self):
        return self.funcCode

    def setFuncCode(self,funcCode):
        self.funcCode = funcCode

    def insertFuncCode(self,funcCode):
        setFocus(self.mainWindow)
        funcBox = self.mainWindow.children()[-3]
        funcColorImg = np.array(funcBox.capture_as_image())
        funcGrayImg = cv.cvtColor(funcColorImg,cv.COLOR_RGB2GRAY)
        _, funcBinaryImg = cv.threshold(funcGrayImg,254,255,cv.THRESH_BINARY)
        locTuple = np.nonzero(funcBinaryImg==255)
        yLoc = np.median(locTuple[0]).astype(np.uint8)
        xLoc = np.median(locTuple[1]).astype(np.uint8)
        funcBox.click_input(coords=(xLoc,yLoc))
        funcBox.type_keys('^a{DELETE}')
        funcBox.type_keys(funcCode+'{ENTER}')
        self.funcWindow = self.app.window(title_re=f".*{re.sub('[A-Z]*','',funcCode)}.*")
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
    window.wait('visible',timeout=5)
