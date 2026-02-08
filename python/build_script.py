import subprocess
import sys
from pathlib import Path
import shutil
import tomllib

def run_quiet(cmd):
  result = subprocess.run(
    cmd,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True
  )
  if result.returncode != 0:
    raise RuntimeError(
      f"Command failed: {' '.join(cmd)}\n{result.stderr}"
    )

def read_version(pyproject: Path) -> str:
  with pyproject.open("rb") as f:
    data = tomllib.load(f)
  return data["project"]["version"]

def wheel_version(dist_dir="dist"):
  for p in Path(dist_dir).glob("*.whl"):
    return p.name.split("-")[1]
  raise RuntimeError("No wheel found in dist/")

# important file paths
repo_root_path = Path(__file__).resolve().parent.parent
dll_file_path = repo_root_path / Path(r"dotnet\SWAutomation.Core\bin\x64\SWAutomation.Core.dll")
dotnet_dir_path = Path(r"src\swauto\_dotnet")
dotnet_file_path = Path(r"src\swauto\_dotnet\SWAutomation.Core.dll")
dist_dir_path = Path(r"dist")
egg_info_path = Path(r"src\swauto.egg-info")
pyproject_path = Path(r"pyproject.toml")

# check if version number has been updated
current_version_num = read_version(pyproject_path)
previous_version_num = wheel_version(str(dist_dir_path))

if current_version_num == previous_version_num:
  print("Update version number before building.")
else:
  print("Version number has been updated, continuing with build...")

  # remove old dll and copy in new one
  try:
    dotnet_file_path.unlink()
  except:
    print("Old dll file already removed.")
  try:
    shutil.copy(str(dll_file_path), str(dotnet_dir_path / "SWAutomation.Core.dll"))
    print("Successfully copied over new dll file.")
  except:
    print("Failed to copy over new dll file.")

  # remove old build artifacts
  try:
    shutil.rmtree(dist_dir_path)
    shutil.rmtree(egg_info_path)
    print("Old build artifacts successfully removed.")
  except:
    print("Failed to remove old build artifacts.")

  # make sure deps are installed in the venv
  try:
    run_quiet([sys.executable, "-m", "pip", "install", "-U", "pip", "build"])
    print("Deps successfully confirmed/installed.")
  except:
    print("Failed to install deps.")

  # build package
  try:
    run_quiet([sys.executable, "-m", "build"])
    print("Package successfully built.")
  except:
    print("Failed to build package.")

  # upload to pypi
  try:
    run_quiet([sys.executable, "-m", "pip", "install", "-U", "twine"])
  except:
    print("Failed to install deps.")
  try:
    run_quiet([sys.executable, "-m", "twine", "upload", str(dist_dir_path / "*")])
    print("Successfully uploaded to PyPI.")
  except:
    print("Failed to upload to PyPI")
