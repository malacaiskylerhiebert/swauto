from __future__ import annotations
import os
import clr # pythonnet

_DLL_DIR = os.path.join(os.path.dirname(__file__), "_dotnet")
_DLL_PATH = os.path.join(_DLL_DIR, "SWAutomation.Core.dll")

os.environ["PATH"] = _DLL_DIR + os.pathsep + os.environ.get("PATH", "")

clr.AddReference(_DLL_PATH)

from SWAutomation.Core import SWConnect

def attach_or_launch(visible: bool = True):
  """
  Returns a SolidWorks SldWorks COM object (via C#)
  """
  return SWConnect.AttachOrLaunch(visible)