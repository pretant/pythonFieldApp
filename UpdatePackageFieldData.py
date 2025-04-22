import subprocess
import time

import requests

exe_name = "C2 Field App.exe"
exe_url = "https://pretant.github.io/packagefielddata/{}".format(exe_name)


def update_software(url):
    response = requests.get(url)
    if response.status_code == 200:
        with open(exe_name, "wb") as f:
            f.write(response.content)
    else:
        print("Could not download update.")


def main():
    print("Waiting for the application to close...")
    time.sleep(5)
    print("Updating...")
    update_software(exe_url)
    print("Update completed. Please restart the application.")
    with open("delete_update_script.bat", "w") as bat_file:
        bat_file.write("timeout 5\n")  # Wait for 5 seconds before deleting
        # Terminate UpdatePackageFieldData.exe process forcefully
        bat_file.write(f"taskkill /IM UpdatePackageFieldData.exe /F\n")
        bat_file.write(f"del UpdatePackageFieldData.exe\n")
        bat_file.write("del %0")  # Delete the batch file itself after completion

    subprocess.Popen("delete_update_script.bat", shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE).wait()


if __name__ == "__main__":
    main()
