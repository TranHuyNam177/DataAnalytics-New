from automation.finance import *
import pyautogui

# Chưa viết xong

def loginFlex():
    w, h = pyautogui.size()
    # Open Flex
    while True:
        windowbuttonImage = join(dirname(__file__),'ImageFlex','WindowButton.png')
        windowbuttonLocation = pyautogui.locateCenterOnScreen(windowbuttonImage,confidence=0.9)
        print(f'windowbuttonLocation = {windowbuttonLocation}')
        if windowbuttonLocation:
            break
    pyautogui.click(windowbuttonLocation)
    pyautogui.sleep(1)
    pyautogui.write('FLEX_PROD',0.1)
    pyautogui.press('enter')
    time.sleep(5)
    # Input username
    while True:
        usernameInputImage = join(dirname(__file__),'ImageFlex','UsernameInput.png')
        initialPoint = pyautogui.locateCenterOnScreen(usernameInputImage)
        if initialPoint:
            break
    usernameLocationx = initialPoint.x + 10
    usernameLocationy = initialPoint.y
    pyautogui.click(usernameLocationx,usernameLocationy)
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('backspace')
    pyautogui.write('1727')
    # Input password
    while True:
        passwordInputImage = join(dirname(__file__),'ImageFlex','PasswordInput.png')
        initialPoint = pyautogui.locateCenterOnScreen(passwordInputImage)
        if initialPoint:
            break
    passwordLocationx = initialPoint.x + 10
    passwordLocationy = initialPoint.y
    pyautogui.click(passwordLocationx,passwordLocationy)
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('backspace')
    pyautogui.write('WY19LP76')
    # Click OK
    okInputImage = join(dirname(__file__),'ImageFlex','OK.png')
    okLocation = pyautogui.locateCenterOnScreen(okInputImage)
    pyautogui.click(okLocation)
