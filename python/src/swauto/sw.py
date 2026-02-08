from __future__ import annotations

import os
import threading
from pathlib import Path
import clr # pythonnet

# function to find repo root folder path
def _find_repo_root(start: Path) -> Path:
  for p in [start, *start.parents]:
    if (p / "pyproject.toml").is_file() or (p / ".git").exists():
      return p
  raise FileNotFoundError("Could not locate repo root (pyproject.tml/.git not found).")

# function to find dotnet dll path
def _get_dotnet_dll_path() -> Path:
  override = os.environ.get("SWAUTO_DOTNET_DIR")
  if override:
    dotnet_dir = Path(override).expanduser().resolve()
  else:
    pkg_dir = Path(__file__).resolve().parent
    repo_root = _find_repo_root(pkg_dir)
    dotnet_dir = (repo_root / "dotnet").resolve()
  
  dll = dotnet_dir / "SWAutomation.Core" / "bin" / "x64" / "SWAutomation.Core.dll"

  if not dotnet_dir.is_dir():
    raise FileNotFoundError(f"dotnet directory not found: {dotnet_dir}")
  if not dll.is_file():
    raise FileNotFoundError(f"DLL not found: {dll}")
  
  dll.relative_to(dotnet_dir)

  return dll

def _load_dotnet():
  if getattr(_load_dotnet, "_loaded", False):
        return
  dll_path = _get_dotnet_dll_path()
  clr.AddReference(str(dll_path))
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
            raise RuntimeError("SwInstance must be used from a single Python thread.")

    def connect(self, visible: bool = True, attach_if_running: bool = True):
        """
        Connect to SolidWorks (attach or launch).
        Idempotent.
        """
        self._check_thread()

        if self._session is not None:
            return

        _load_dotnet()
        from SWAutomation.Core import SwSession  # type: ignore

        self._session = SwSession()
        self._session.Connect(visible, attach_if_running)

    def open_assembly(self, path: str, silent: bool = True):
        """
        Open an assembly document.
        """
        self._check_thread()

        if self._session is None:
            raise RuntimeError("Call connect() first.")

        return self._session.OpenAssembly(path, silent)