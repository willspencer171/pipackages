import sys
import os
import time
import pandas as pd
import subprocess
import re

# Let's update my packages
# subprocess.getoutput calls the command line command and stores its output as string
outdates = subprocess.getoutput([sys.executable, "-m", "pip", "list", "-o"])
outdates = outdates.split("\n")[2:]

# Get the name of the package and upgrade it using subprocess.call
for index, package in enumerate(outdates):
    try:
        package = re.split(" ", package)
        subprocess.call([sys.executable, "-m", "pip", "install", "--upgrade", package[0]])
        print(f"Package {index + 1}: '{package[0]}' upgraded")
    except OSError as e:
        print("ENCOUNTERED AN ERROR")
        print(e)

# Write updated outdates file - in case of issues with upgrading etc
# json format used as freeze format no longer supported with outdated list
output_file = os.path.dirname(sys.executable) + "outdates.txt"
output_data = subprocess.getoutput([sys.executable, "-m", "pip", "list", "-o", "format=json", "| jq -r '.[]", "| .name+\"==\"+.latest_version'"])
with open(output_file, "w") as f:
    f.write(output_data)

if not len(output_data):
    print("All installed packages are up to date :)")

# Replace any None values with empty string
def as_text(value):
    if value is None:
        return ""
    return str(value)

file_source = "./Python Packages.xlsx"

# Import metadata module dependent on Python version info,
# For back compat
if sys.version_info >= (3, 8):
    from importlib import metadata as importlib_metadata
else:
    import importlib_metadata

# Obtain Package distributions
dists = importlib_metadata.distributions()
dists_list = []

# Find path, version, name and last update time of distribution
for dist in dists:
    path = str(dist.locate_file(dist.files[0])).replace(str(dist.files[0]), "")
    version = dist.version
    name = dist.name
    pack_time = time.ctime(os.path.getctime(path))
    pack_time = pack_time.split(" ")
    pack_time = [x for x in pack_time if x != ""]
    pack_time.pop(3)
    pack_time = " ".join(pack_time)
    dists_list.append([name, version, pack_time])

df = pd.DataFrame(dists_list,
                    columns=["Package Name", "Version", "Last Updated"])
df.set_index("Package Name", inplace=True)
print(df)

with pd.ExcelWriter(file_source, engine="openpyxl") as xlsx:
    sheet_name="Packages"
    df.to_excel(xlsx, sheet_name)

    # set column width
    ws = xlsx.sheets[sheet_name]
    for column_cells in ws.columns:
        length = max(len(as_text(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 1
