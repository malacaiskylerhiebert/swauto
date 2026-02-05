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
  Attach to a running SolidWorks session or launch a new one.

  Parameters
  ----------
  visible:
      If True, shows the SolidWorks UI.

  Returns
  -------
  SldWorks
      SolidWorks application COM object.

  Notes
  -----
  Requires Windows, 64-bit Python, SolidWorks installed, and pythonnet.
  """
  return SWConnect.AttachOrLaunch(visible)