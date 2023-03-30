import subprocess
import pathlib

# this will execute the installation of requirement packages
subprocess.call(["pip", "install", "-r", str(pathlib.Path(__file__).parent.resolve()) + "/requirements.txt"])
