import ctypes
import subprocess

import addonHandler
import globalPluginHandler
from scriptHandler import script

addonHandler.initTranslation()

kernel32 = ctypes.windll.kernel32


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = 'Open CMD'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    @script(
        description=_('Open CMD'),
        gesture='kb:nvda+control+tab'
    )
    def script_open_cmd(self, gesture):
        out = ctypes.c_void_p()
        if not kernel32.Wow64DisableWow64FsRedirection(ctypes.byref(out)):
            message('Failed.')
            return
        try:
            subprocess.Popen('cmd.exe /s /k pushd "%userprofile%"')
        finally:
            kernel32.Wow64RevertWow64FsRedirection(ctypes.byref(out))
