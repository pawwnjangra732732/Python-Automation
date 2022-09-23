import os
import pywinauto
from pywinauto.keyboard import send_keys as sk
from pywinauto import Application
import os
import time
import datetime
#varibles Declaring
pdf_download_path=f"C:\\Users\\{os.getlogin()}\\Downloads\\"
app=Application(backend="win32").connect(title_re="Save As",timeout=15)
dlg=app.top_window()
address = app.SaveAs.child_window(title="Address band toolbar",class_name="ToolbarWindow32").wrapper_object()
print(dlg.print_control_identifiers())
address.click()
time.sleep(2)
dlg.set_focus()
sk("{BACKSPACE}"*40)
dlg.set_focus()
sk(pdf_download_path+"{ENTER}")
time.sleep(1)
filename="Cordial Care Statement "+datetime.date.today().strftime("%m%d%Y")
app.SaveAs.Edit1.type_keys(filename+"{ENTER}",with_spaces=True)
try:
    app = Application(backend="win32").connect(title="Confirm Save As", timeout=10)
    dlg = app['Confirm Save As']
    dlg.set_focus()
    yesbtn = app.ConfirmSaveAs.child_window(title="&Yes", class_name="Button").wrapper_object()
    yesbtn.click()
except:
    pass
