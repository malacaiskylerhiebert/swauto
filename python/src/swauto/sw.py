from __future__ import annotations

import os
import threading
from pathlib import Path
import clr # pythonnet
import math

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

  def open_part(self, path: str, silent: bool = True) -> Part:
    self._check_thread()
    if self._session is None:
      raise RuntimeError("Call connect() first.")
    doc_id = self._session.OpenPart(path, silent)
    return Part(self._session, str(doc_id))

  def open_assembly(self, path: str, silent: bool = True) -> Assembly:
    """
    Open an assembly document.
    """
    self._check_thread()

    if self._session is None:
      raise RuntimeError("Call connect() first.")

    doc_id = self._session.OpenAssembly(path, silent)
    return Assembly(self._session, str(doc_id))
  
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
  
class _Document:
  def __init__(self, session, doc_id: str):
    self._session = session
    self._id = doc_id

  def save(self, silent: bool = True):
    self._session.SaveDocument(self._id, bool(silent))
  
  def save_as(self, out_path: str, silent: bool = True):
    self._session.SaveAsDocument(self._id, str(out_path), bool(silent))

  def rebuild(self, top_only: bool = False):
    self._session.RebuildDocument(self._id, bool(top_only))

  def close(self, save: bool = False, silent_save: bool = True):
    self._session.CloseDocument(self._id, bool(save), bool(silent_save))
  
class Part(_Document):
  pass

class Assembly(_Document):
  def add_component(self, part, x: float = 0.0, y: float = 0.0, z: float = 0.0):
    if type(part) == str:
      part = part
    elif type(part) == Part:
      part = part._id
    self._session.AssemblyAddComponent(self._id, part, x, y, z)

  def remove_component(self, comp):
    """
    comp may be:
      - component name (str)
      - Part object (uses its _id)
      - docId string
    """
    comp_ref = comp._id if hasattr(comp, "_id") else str(comp)
    self._session.AssemblyRemoveComponent(self._id, comp_ref)

  def set_fixed(self, comp, fixed: bool):
    comp_ref = comp._id if hasattr(comp, "_id") else str(comp)
    self._session.AssemblySetComponentFixed(self._id, comp_ref, bool(fixed))

  def translate_component(self, comp, dx: float, dy: float, dz: float):
    """dx/dy/dz in meters."""
    comp_ref = comp._id if hasattr(comp, "_id") else str(comp)
    self._session.AssemblyTranslateComponent(
      self._id, 
      comp_ref, 
      float(dx), 
      float(dy), 
      float(dz)
    )

  def rotate_component(self, comp, rx: float, ry: float, rz: float):
    """
    Rotate component by rx, ry, rz radians about X, Y, Z.
    """
    comp_ref = comp._id if hasattr(comp, "_id") else str(comp)
    self._session.AssemblyRotateComponent(
      self._id,
      comp_ref,
      float(rx),
      float(ry),
      float(rz),
    )