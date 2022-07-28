from pywinauto.application import Application

PATH = r'C:\PHS-APP\PROD\FLEX_PROD\@DIRECT.exe'
app = Application().start(PATH)
# Login Flex
user_name = app['Login'].child_window(auto_id='txtUserName').click()
user_name.type_keys('1964', with_spaces=False)
pwd = app['Login'].child_window(auto_id='txtPassword').click()
pwd.type_keys('hmWn7XEw', with_spaces=False)
app['Login'].child_window(auto_id='btnOK').click()
# Thao tác tự động trong app Flex
app['.:: Flex 6.8.17 On 28/06/2022 ::.'].print_control_identifiers()
