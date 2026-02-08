from __future__ import annotations

import os
import threading
from pathlib import Path
import clr # pythonnet

def _resolve_dotnet_dll() -> Path:
  # sw.py is in swauto/, so _dotnet is sibling folder
  dll = Path(__file__).resolve().parent / "_dotnet" / "SWAutomation.Core.dll"
  if not dll.is_file():
    raise FileNotFoundError(
      f"Missing SWAutomation.Core.dll at {dll}. "
      "Ensure the DLL is copied into swauto/_dotnet/."
    )
  return dll


def _load_dotnet():
  if getattr(_load_dotnet, "_loaded", False):
    return
  clr.AddReference(str(_resolve_dotnet_dll()))
  _load_dotnet._loaded = True

## future commands
# - connect
# - open_assembly
# - rebuild
# - exportstep
# - closedoc

class SWInstance:
  def __init__(self):
    self._owner_thread = threading.get_ident()
    self._session = None

  def _check_thread(self):
    if threading.get_ident() != self._owner_thread:
      raise RuntimeError("SWInstance must be used from a single Python thread.")

  def connect(self, visible: bool = True, attach_if_running: bool = True):
    """
    Connect to SolidWorks (attach or launch).
    Idempotent.
    """
    self._check_thread()

    if self._session is not None:
      return

    _load_dotnet()
    from SWAutomation.Core import SWSession  # type: ignore

    self._session = SWSession()
    self._session.Connect(visible, attach_if_running)

  def open_assembly(self, path: str, silent: bool = True):
    """
    Open an assembly document.
    """
    self._check_thread()

    if self._session is None:
      raise RuntimeError("Call connect() first.")

    return self._session.OpenAssembly(path, silent)
  
  def get_revision(self):
    """
    Obtains SolidWorks revision number.
    """
    self._check_thread()

    if self._session is None:
      raise RuntimeError("Call connect() first.")
    
    return self._session.GetRevision()
    
  def close(self):
    """
    Closes the instance and only closes the program if it was opened by swauto.
    """
    self._check_thread()

    if self._session is None:
      raise RuntimeError("Call connect() first.")
    
    return self._session.Shutdown(force=False)
  
  def kill(self):
    """
    Kills all instances of SolidWorks.
    """
    self._check_thread()

    if self._session is None:
      raise RuntimeError("Call connect() first.")
    
    return self._session.Shutdown(force=True)