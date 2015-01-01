
'''
The .msi resulting from `setup.py bdist_msi` contains some errors. This script fixes them, but will only support an "everyone" "ALLUSERS=1" install.

Caveat: only tested with Python 2.7 win32
TODO: remove UI where user picks between everyone install or "just for me"

Build with:
cd pywin32
setup.py bdist_msi
pywin32_msifix.py dist/pywin32-219.0.win32-py2.7.msi
Install with: 
msiexec /I pywin32-219.0.win32-py2.7.msi /qn /L* pywin32.log ALLUSERS=1
'''

import msilib
import sys

if len(sys.argv) == 2:
    msifile = sys.argv[1]
else:
    msifile = 'dist/pywin32-219.0.win32-py2.7.msi'

#import shutil
#shutil.copyfile('dist/pywin32-219.0.win32-py2.7.msi - Copy.msi', 'dist/pywin32-219.0.win32-py2.7.msi')

db = msilib.OpenDatabase(msifile, msilib.MSIDBOPEN_DIRECT)

def iterateView(query, record=None):
    view = db.OpenView(query)
    view.Execute(record)
    while True:
        try:
            record = view.Fetch()
        except msilib.MSIError as e:
            if e.message != 'unknown error 103':
                raise
            view.Close()
            return
        yield (view, record)

#InstallExecuteSequence: remove .py post-install scripts (install_script.2.7, install_script.X)
for view, record in iterateView("select * from InstallExecuteSequence"):
    #print record.GetString(1)
    if record.GetString(1) in ('install_script.2.7', 'install_script.X'):
        view.Modify(msilib.MSIMODIFY_DELETE, record)
# add SetDLLDirToSystem32=712.
view = db.OpenView("select * from InstallExecuteSequence")
view.Execute(None)
record = msilib.CreateRecord(3)
record.SetString(1, 'SetDLLDirToSystem32')
record.SetInteger(3, 712)
view.Modify(msilib.MSIMODIFY_INSERT, record)
view.Close()

#CustomAction: add SetDLLDirToSystem32,51,pywin32_system32,[SystemFolder]. 
view = db.OpenView("select * from CustomAction")
view.Execute(None)
record = msilib.CreateRecord(4)
record.SetString(1, 'SetDLLDirToSystem32')
record.SetInteger(2, 51)
record.SetString(3, 'pywin32_system32')
record.SetString(4, '[SystemFolder]')
view.Modify(msilib.MSIMODIFY_INSERT, record)
view.Close()

#Directory: change Lib and Scripts Directory_parent to TARGETDIR2.7
for view, record in iterateView("select * from Directory"):
    #print record.GetString(1)
    if record.GetString(1) in ('Lib', 'Scripts'):
        record.SetString(2, 'TARGETDIR2.7')
        view.Modify(msilib.MSIMODIFY_REPLACE, record)

db.Commit()
