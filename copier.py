

import os
import shutil
import vss
from pyuac import main_requires_admin


@main_requires_admin(return_output=False)
def main():
    file_export = "./powerpoint/defense_export.pptx"
    file = "./powerpoint/defense.pptx"
    

    # shutil.copyfile(file, file_export)

    full_path = os.path.abspath(file)


    local_drives = set()
    local_drives.add('C')
    # Initialize the Shadow Copies
    sc = vss.ShadowCopy(local_drives)
    # An open and locked file we want to read
    shadow_path = sc.shadow_path(full_path)

    print(shadow_path)

    shutil.copy(shadow_path, file_export)

    print('testing')


if __name__ == "__main__":
    rv = main()
    # if not rv:
    #     print("I must have already been Admin!")
    # else:
    #     admin_stdout_str, admin_stderr_str, *_ = rv
    #     if "Do stuff" in admin_stdout_str:
    #         print("It worked.")
