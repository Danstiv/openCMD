import ctypes
import subprocess

import addonHandler
import globalPluginHandler
import ui
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
        subprocess.Popen('cmd.exe /s /k pushd "%userprofile%"')
