import collections
import collections.abc
import pptx
import win32com.client as client
import os
import shutil
import vss
from pyuac import main_requires_admin
import time
import subprocess



def prev_slide(number):
    if number > 0:
        return number - 1
    else:
        return 0


def next_slide(number, ln):
    if number < ln - 1:
        return number + 1
    else:
        return ln - 1
    

def create_svelte_route(i, ln, routes_directory):
    # I should add a checker here that looks in the route directory
    # if its there and does not overwrite it if theres some keyfile 
    # indicating dynamic content

    folder_name = f"slide_{i}"
    folder_path = os.path.join(routes_directory, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    # Create a text file in the folder
    svelte_path = os.path.join(folder_path, "+page.svelte")
    with open(svelte_path, "w") as f:
        f.write(
            f"""
<script>
  import {{ onMount }} from 'svelte';
  import {{ writable }} from 'svelte/store';
  // <img src="/src/slides_png/slide_{i}.png" alt="slide_{i}" width="1000" height="500">

  const width = writable(0);
  const height = writable(0);
  onMount(() => {{
    width.set(window.innerWidth);
    height.set(window.innerHeight);
    document.addEventListener('keydown', function(event) {{
      //Check if the pressed key is the left arrow key
      if (event.key === 'ArrowLeft') {{
        // Navigate to the desired URL when the left arrow key is pressed
        window.location.href = '/slide_{prev_slide(i)}'; // Replace with your desired URL
      }}
      //Check if the pressed key is the right arrow key
      else if (event.key === 'ArrowRight') {{
        // Navigate to the desired URL when the right arrow key is pressed
        window.location.href = '/slide_{next_slide(i, ln)}'; // Replace with your desired URL
      }}
    }});
  }});
</script>


<img src="/src/slides_png/slide_{i}.png" alt="slide_{i}" width="{{$width}}" height="{{$height}}">
              
"""
        )

    js_path = os.path.join(folder_path, "+page.js")
    with open(js_path, "w") as f:
        f.write(
            """
// since there's no dynamic data here, we can prerender
// it so that it gets served as a static asset in production
export const prerender = true;
"""
        )
        # end

# @main_requires_admin(return_output=False)
# def main():
#     file_export = "./powerpoint/defense_export.pptx"
#     file = "./powerpoint/defense.pptx"
    

#     # shutil.copyfile(file, file_export)

#     full_path = os.path.abspath(file)


#     local_drives = set()
#     local_drives.add('C')
#     # Initialize the Shadow Copies
#     sc = vss.ShadowCopy(local_drives)
#     # An open and locked file we want to read
#     shadow_path = sc.shadow_path(full_path)

#     print(shadow_path)

#     shutil.copy(shadow_path, file_export)

#     print('testing')


def run_copy():
    subprocess.run(["python", "copier.py"])

def finish_up():
    run_copy()
    file_export = "./powerpoint/defense_export.pptx"
    export_path = ".\\web\\defense\\src\\slides_png\\"
    full_export_path = os.path.abspath(export_path)

    full_file_path = os.path.abspath(file_export)


    # works
    # ppt=Presentation("/Users/Andrew/defense.pptx")
    ppt = pptx.Presentation(file_export)

    notes = []



    # https://stackoverflow.com/questions/61815883/python-pptx-export-img-png-jpeg


    Application = client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Open(full_file_path)





    # Define the base directory where you want to create the folders
    routes_directory = "./web/defense/src/routes/"

    # Define the number of folders you want to create
    # num_folders = 5

    # Loop through the range of folder numbers and create each folder
    for i, (note_slide, img_slide) in enumerate(zip(ppt.slides, Presentation.Slides)):
        textNote = note_slide.notes_slide.notes_text_frame.text
        notes.append((i, textNote))

        if "dyno" not in textNote:
            img_slide.Export(f"{full_export_path}\\slide_{i}.png", "PNG")

            create_svelte_route(i, len(ppt.slides), routes_directory)

        else:
            print("found dynamic slide")

    print("done?")
    # Application.Quit()
    Presentation.Close()
    Presentation = None
    Application = None


if __name__ == "__main__":

    finish_up()

