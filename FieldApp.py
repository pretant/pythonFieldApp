import math
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.messagebox as messagebox
import traceback
import webbrowser
from datetime import datetime
from tkinter import filedialog, StringVar, simpledialog, ttk
from zipfile import ZipFile, ZIP_DEFLATED

import customtkinter as ctk
import exifread
import geopy.distance
import pandas as pd
import psutil
import pyperclip
import requests
from PIL import Image, ExifTags
from thefuzz import process, fuzz
from tkcalendar import DateEntry

exe_name = "C2 Field App.exe"
version_url = "https://pretant.github.io/packagefielddata/version.txt"
version_history_url = "https://pretant.github.io/packagefielddata/versionhistory/"


def print_to_widget(text, newline=True, color='white', url=None):
    text_space.configure(state='normal')

    # Insert the text with newline if requested
    insert_position = text_space.index("end-1c")
    text_to_insert = str(text) + ('\n' if newline else '')
    text_space.insert("end", text_to_insert)

    # Calculate the end position of the inserted text for color or hyperlink application
    end_position = text_space.index(f"{insert_position}+{len(text_to_insert)}c")

    # Configure tag for color if it hasn't been configured yet in a global context
    if color and color not in text_space.tag_names():
        text_space.tag_config(color, foreground=color)

    # Apply the color tag to the inserted text
    text_space.tag_add(color, insert_position, end_position)

    if url:
        # Define a unique tag name based on the current end position (to ensure uniqueness)
        hyperlink_tag = f"hyperlink-{insert_position.replace('.', '-')}"
        # Add the tag to the inserted text for hyperlink
        text_space.tag_add(hyperlink_tag, insert_position, end_position)
        # Configure the tag with hyperlink styles and bind a click event
        text_space.tag_config(hyperlink_tag, foreground="#87CEEB", underline=True)
        text_space.tag_bind(hyperlink_tag, "<Button-1>", lambda e, url=url: webbrowser.open_new(url))

    text_space.configure(state='disabled')
    text_space.see(tk.END)


def display_exception():
    exc_type, exc_value, exc_traceback = sys.exc_info()
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))

    # Calculate text_widget dimensions based on the error message
    lines = error_msg.splitlines()
    max_width = max(len(line) for line in lines)
    num_lines = len(lines)

    error_window = tk.Toplevel()
    error_window.configure(bg="#18191A")
    error_window.title("Error")

    label = ttk.Label(error_window, text="An unexpected error occurred:", style="DarkTheme.TLabel", anchor="w")
    label.pack(padx=10, pady=10, anchor="w")

    text_widget = tk.Text(error_window, wrap=tk.NONE, padx=10, pady=10, bg="#18191A", fg="#B3B3B3",
                          insertbackground="#B3B3B3", selectbackground="#7a7a7a",
                          width=max_width, height=num_lines)
    text_widget.insert(tk.END, error_msg)
    text_widget.config(state=tk.DISABLED)  # Prevent print_text editing
    text_widget.pack(expand=True, fill=tk.BOTH)

    def copy_error_message():
        error_window.clipboard_clear()
        error_window.clipboard_append(error_msg)
        tooltip.show_tip()
        error_window.after(1500, tooltip.hide_tip)

    copy_exception_button = ttk.Button(error_window, text="Copy", command=copy_error_message, style="DarkTheme.TButton")
    copy_exception_button.pack(pady=10)

    tooltip = ToolTip(copy_exception_button, "Copied to clipboard")


def get_current_version():
    return "2.2.1"


def open_version_history(event):
    webbrowser.open(version_history_url)


def get_latest_version():
    try:
        response = requests.get(version_url)
    except requests.exceptions.RequestException:
        print_to_widget("Failed to fetch latest version.")
        return None

    if response.status_code == 200:
        return response.text.strip()
    else:
        print_to_widget(f"Failed to fetch latest version.")
        return None


def start_update_script():
    response = requests.get("https://pretant.github.io/packagefielddata/UpdatePackageFieldData.exe")
    if response.status_code == 200:
        with open("UpdatePackageFieldData.exe", "wb") as f:
            f.write(response.content)
        return subprocess.Popen(["UpdatePackageFieldData.exe"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        print_to_widget("Could not download update script.\n")
        return None


def check_for_updates():
    current_version = get_current_version()
    latest_version = get_latest_version()

    if latest_version:
        if latest_version != current_version:
            print_to_widget(f"New version available: ", newline=False)
            print_to_widget(f"v{latest_version}", color="#87CEEB", url=version_history_url)
            print_to_widget("Do you want to update?")
            response = messagebox.askyesno("Update Available",
                                           f"Do you want to update?")
            if response:
                print_to_widget("Downloading update...")
                for proc in psutil.process_iter():
                    try:
                        if proc.name() == exe_name:
                            print_to_widget("Closing app...")
                            start_update_script()
                            time.sleep(5)  # Give the update script some time to start before killing the main app
                            try:
                                proc.kill()
                            except Exception as e:
                                print_to_widget(f"{e}\n\nClose the app manually to proceed with the update.")
                                pass
                            break
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        pass
            else:
                print_to_widget("\nUpdate cancelled.\n")
        else:
            print_to_widget(f"You are running the latest version.\nClick ", newline=False)
            print_to_widget(f"here", color="#87CEEB", url=version_history_url, newline=False)
            print_to_widget(f" to see version history.\n")
    else:
        print_to_widget("Unable to check for updates.\n")


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.tip_window = None
        self.text = text

    def show_tip(self):
        if self.tip_window or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 30
        if y + 50 > self.widget.winfo_screenheight():
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        if x + 180 > self.widget.winfo_screenwidth():
            x = self.widget.winfo_rootx() - 180
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#242526", foreground="#B3B3B3",
                         relief=tk.SOLID, borderwidth=1, font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


def create_tooltip(widget, text):
    tool_tip = ToolTip(widget, text)

    def enter(event):
        tool_tip.show_tip()

    def leave(event):
        tool_tip.hide_tip()

    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)


# Function to get "date taken" metadata of an image.
def get_date_taken(file_path):
    image_name = os.path.basename(file_path)
    with open(file_path, 'rb') as f:
        tags = exifread.process_file(f, stop_tag='EXIF DateTimeOriginal')
        if 'EXIF DateTimeOriginal' in tags:
            date_taken = tags['EXIF DateTimeOriginal'].values
            date_taken = datetime.strptime(date_taken, '%Y:%m:%d %H:%M:%S').strftime('%m.%d.%Y')
        else:
            date_taken = None
            date_str_queue = queue.Queue()

            def ask_date():
                date_input = simpledialog.askstring("Enter Date",
                                                    f"Unable to find date taken metadata for {image_name}."
                                                    f" Please enter the flight date for this structure (YYYYMMDD)",
                                                    parent=root)
                date_str_queue.put(date_input)

            while date_taken is None:
                root.after(0, ask_date)
                date_str = date_str_queue.get()  # This will block until a result is available
                try:
                    date_taken = datetime.strptime(date_str, '%Y%m%d').strftime('%m.%d.%Y')
                except ValueError:
                    print_to_widget("Invalid date format, please try again.")
                    continue
    return date_taken


# Function to extract GPS data from images using PIL library
def get_gps_from_image(filepath):
    try:
        with Image.open(filepath) as img:
            # Get EXIF data from the image
            exif_data = img._getexif()
            if exif_data:
                # Extract only necessary EXIF tags using PIL's TAGS dictionary
                exif_data = {
                    ExifTags.TAGS[k]: v
                    for k, v in exif_data.items()
                    if k in ExifTags.TAGS
                }
                # Extract GPSInfo tag and its corresponding sub-tags using PIL's GPSTAGS dictionary
                gps_info = exif_data.get('GPSInfo', {})
                if gps_info:
                    gps_info = {
                        ExifTags.GPSTAGS.get(key, key): value
                        for key, value in gps_info.items()
                    }
                    # Extract latitude and longitude data from GPSInfo
                    lat = gps_info.get('GPSLatitude')
                    lat_ref = gps_info.get('GPSLatitudeRef', 'N')
                    lng = gps_info.get('GPSLongitude')
                    lng_ref = gps_info.get('GPSLongitudeRef', 'E')
                    if lat and lng:
                        # Convert latitude and longitude from degrees, minutes, seconds to decimal degrees
                        lat_decimal = lat[0].numerator / lat[0].denominator + (
                                lat[1].numerator / lat[1].denominator) / 60 + (
                                              lat[2].numerator / lat[2].denominator) / 3600
                        lng_decimal = lng[0].numerator / lng[0].denominator + (
                                lng[1].numerator / lng[1].denominator) / 60 + (
                                              lng[2].numerator / lng[2].denominator) / 3600
                        # Apply negative sign to latitude and/or longitude if necessary
                        if lat_ref == 'S':
                            lat_decimal *= -1
                        if lng_ref == 'W':
                            lng_decimal *= -1
                        # Return latitude and longitude as a tuple
                        return lat_decimal, lng_decimal
    except (OSError, AttributeError, KeyError):
        pass
    # Return None if GPS data is not found in the image or there is an error extracting the data
    return None


# Function to calculate distance between two coordinates in feet
def distance_calculator(coord1, coord2):
    if coord1 is None or coord2 is None:
        print_to_widget("   - Unable to calculate distance from GIS. GIS coordinates missing.", color='red')

    # Check if coordinates are strings
    if any(isinstance(c, str) for c in coord1) or any(isinstance(c, str) for c in coord2):
        return None

    if any(math.isnan(c) for c in coord1) or any(math.isnan(c) for c in coord2):
        print_to_widget("   - Unable to calculate distance from GIS. GIS coordinates missing.", color='red')
        return None

    lat1, lon1 = coord1
    lat2, lon2 = coord2

    if not (-90 <= lat1 <= 90) or not (-90 <= lat2 <= 90):
        print_to_widget(f"Invalid latitude value(s): {lat1}, {lat2}")
        return None

    if not (-180 <= lon1 <= 180) or not (-180 <= lon2 <= 180):
        print_to_widget(f"Invalid longitude value(s): {lon1}, {lon2}")
        return None

    # Calculate distance and round to two decimal places
    distance = round(geopy.distance.distance(coord1, coord2).feet, 2)
    return distance


# Function to get the farthest image from the nadir
def get_farthest_from_nadir(root_directory):
    farthest_distances = {}  # Create a dictionary to store the farthest distances for each subfolder

    # Walk through the directory structure
    for dirpath, dirnames, filenames in os.walk(root_directory):
        n_image_path = None
        image_paths = []

        # Iterate through the files in the current directory
        for file in filenames:
            file_path = os.path.join(dirpath, file)

            if file_path.lower().endswith("n.jpg"):
                n_image_path = file_path
            elif file.lower().endswith((".jpg", ".jpeg")):
                image_paths.append(file_path)

        if n_image_path:
            n_coord = get_gps_from_image(n_image_path)

            if n_coord:
                max_distance = 0

                for img_path in image_paths:
                    coord = get_gps_from_image(img_path)
                    if coord:
                        dist_feet = distance_calculator(n_coord, coord)
                        if dist_feet > max_distance:
                            max_distance = dist_feet
                    else:
                        img_name = os.path.basename(img_path)
                        print_to_widget(
                            f"Warning: No GPS data found on {img_name} from structure {os.path.basename(dirpath)}.")
                # Store the farthest distance in the dictionary
                farthest_distances[os.path.basename(dirpath)] = max_distance

            else:
                print_to_widget(f"Warning: No GPS data found on the nadir of {os.path.basename(dirpath)}.")
                farthest_distances[os.path.basename(dirpath)] = None
        elif image_paths:  # Check if the folder had images but no nadir
            print_to_widget(f"Warning: No nadir file found in {os.path.basename(dirpath)}.")
            farthest_distances[os.path.basename(dirpath)] = None

    return farthest_distances


def generate_txt_file(directory_path, pilot_id):
    """
    Writes the names of all subfolders in the given directory to a file named '[pilot_id]_structure_list.txt'.

    :param directory_path: String path to the directory where subfolders are located
    :param pilot_id: The pilot ID to be used in the generated file name
    """
    # Validate if the path exists
    if not os.path.isdir(directory_path):
        print_to_widget(f"Error: The directory '{directory_path}' does not exist or is not a valid directory.")
        return

    # Create (or overwrite) '[pilot_id]_structure_list.txt'
    output_file_path = os.path.join(directory_path, f"{pilot_id}_structure_list.txt")
    with open(output_file_path, 'w', encoding='utf-8') as f:
        # List everything in the directory
        for item in os.listdir(directory_path):
            # Skip "EZPolesForTrans"
            if item == "EZPolesForTrans":
                continue
            # Build the full path of the item
            item_path = os.path.join(directory_path, item)
            # Check if the item is a directory
            if os.path.isdir(item_path):
                f.write(item + '\n')

    print_to_widget(f"\n{pilot_id}_structure_list.txt has been generated")


def get_directory_size(directory):
    total_size = 0
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.exists(fp):
                total_size += os.path.getsize(fp)
    return total_size


def update_progress_bar(progress_queue):
    if not progress_queue.empty():
        progress = progress_queue.get()
        progress_bar_canvas.coords(progress_bar_rect, 0, 0, progress * 4.48, 20)  # Update the progress bar
        progress_bar_canvas.itemconfig(progress_bar_percentage, text=f"{int(progress)}%")  # Update the percentage
        root.update_idletasks()
    if not progress_queue.empty():
        root.after(1000, update_progress_bar, progress_queue)


# Initialize cancel_zip flag to False (used to cancel the zipping process)
cancel_zip = False


def request_cancel():
    global cancel_zip
    cancel_zip = True
    print_to_widget("Cancellation requested. Exiting zipping process.")


def zip_directory(source, destination, progress_queue):
    global cancel_zip
    total_files = sum(len(files) for _, _, files in os.walk(source))
    zipped_files = 0

    base_folder_name = os.path.basename(source)

    with ZipFile(destination, 'w', ZIP_DEFLATED) as zf:
        for directory, _, files in os.walk(source):
            for file in files:
                # Check if the cancel_zip flag is set to True (when cancel button is pressed)
                if cancel_zip:
                    print(123)
                    return
                filepath = os.path.join(directory, file)
                arc_name = os.path.join(base_folder_name, os.path.relpath(filepath, source))
                zf.write(filepath, arc_name)
                zipped_files += 1
                # Update the progress queue with the progress percentage
                progress_percentage = (zipped_files / total_files) * 100
                progress_queue.put(progress_percentage)


def packaging_thread_function():
    # Check if the required information is provided
    if not dir_path.get() or not team_number_entry.get() or not date_entry.get_date():
        messagebox.showerror("Error", "Please input directory, flight date, and team number.")
        return

    progress_queue = queue.Queue()

    # Create a separate thread for the packaging process
    packaging_thread = threading.Thread(target=rename_and_zip_directory, args=(progress_queue,))
    packaging_thread.start()


def upload_trans_data():
    webbrowser.open('https://c2groupoffice-my.sharepoint.com/:f:/r/personal/c2drone_c2groupoffice_onmicrosoft_com'
                    '/Documents/UAV%20Projects/SCE/Field%20Uploads/2025/Transmission?csf=1&web=1&e=6CyMUX')


def upload_distro_data():
    webbrowser.open('https://c2groupoffice-my.sharepoint.com/:f:/r/personal/c2drone_c2groupoffice_onmicrosoft_com'
                    '/Documents/UAV%20Projects/SCE/Field%20Uploads/2025/Distribution?csf=1&web=1&e=a6eaW5')


def find_innermost_folders(directory):
    innermost_folders = []
    for dirpath, dirnames, files in os.walk(directory):
        # If 'dirnames' is empty, it means 'dirpath' is an innermost folder
        if not dirnames and files:
            innermost_folders.append(dirpath)
    return innermost_folders


# Define a function to find the closest match for a folder name
def find_closest_match(folder_name, folder_path, choices, dataframe, verbose=True, resolve=False, no_dist_issue=True):
    global closest_distance
    # If the folder name is not a valid structure ID, proceed with the following steps
    # Find the image that ends with n.jpg
    for file in os.listdir(folder_path):
        if file.lower().endswith("n.jpg"):
            image_path = os.path.join(folder_path, file)
            break
    else:
        if no_dist_issue:
            # If no image found, use fuzzywuzzy to find the closest match
            closest_match = process.extractOne(folder_name, choices)[0]
            choices.remove(closest_match)  # Remove the matched option
            if verbose:
                print_to_widget(f"   - No nadir image found.")
                print_to_widget(f"   - Closest ID match: {closest_match}")
            return folder_name, closest_match, choices
        else:  # Use first image for image_path
            image_path = os.path.join(folder_path, os.listdir(folder_path)[0])

    # Get GPS coordinates from the image
    lat, lon = get_gps_from_image(image_path)
    if lat is None or lon is None:
        if no_dist_issue:
            closest_match = process.extractOne(folder_name, choices)[0]
            choices.remove(closest_match)  # Remove the matched option
            if verbose:
                print_to_widget(f"   - Nadir image does not have gps data.")
                print_to_widget(f"   - Closest ID match: {closest_match}")
            return folder_name, closest_match, choices
        else:
            pass

    # Calculate the distance between the image's coordinates and each structure coordinate in the date_taken_df
    distances = []
    for dataframe_index, dataframe_row in dataframe.iterrows():
        if pd.isna(dataframe_row['Latitude']) or pd.isna(dataframe_row['Longitude']):
            continue
        mapped_lat, mapped_lon = dataframe_row["Latitude"], dataframe_row["Longitude"]
        distance = distance_calculator((lat, lon), (mapped_lat, mapped_lon))
        distances.append((dataframe_row["Structure ID"], distance))

    # Find the closest distance and the associated structure ID
    if distances:
        closest_match = min(distances, key=lambda x: x[1])[0]
        closest_distance = min(distances, key=lambda x: x[1])[1]
        if resolve:
            choices.remove(closest_match)  # Remove the matched option
        if verbose:
            print_to_widget(f"   - Closest nadir match: {closest_match}.")
            if closest_distance >= 500:
                print_to_widget(f"   - Distance from nadir: ", newline=False)
                print_to_widget(f"{closest_distance} feet", color='red', newline=False)
                print_to_widget(f".")
            if closest_distance >= 150:
                print_to_widget(f"   - Distance from nadir: ", newline=False)
                print_to_widget(f"{closest_distance} feet", color='#FFA500', newline=False)
                print_to_widget(f".")
            else:  # If the distance is less than 200 feet
                print_to_widget(f"   - Distance from nadir: ", newline=False)
                print_to_widget(f"{closest_distance} feet.")
        return folder_name, closest_match, choices


def resolve_duplicates(matches, available_choices, structure_dict, dataframe, no_dist_issue=True, verbose=True):
    resolved_duplicates = {}
    while any(isinstance(value, list) for value in matches.values()):
        duplicates = {key: value for key, value in matches.items() if isinstance(value, list)}
        new_matches = {key: value for key, value in matches.items() if not isinstance(value, list)}

        for match, folders in duplicates.items():
            # Compute scores for all folders against the match
            scores = {folder: fuzz.ratio(folder, match) for folder in folders}

            # Identify the most accurate folder (highest score)
            most_accurate_folder = max(scores, key=scores.get)
            new_matches[match] = most_accurate_folder
            resolved_duplicates[match] = most_accurate_folder
            try:
                available_choices.remove(match)
            except (ValueError, KeyError):
                pass

            # All other folders will use find_closest_match
            other_folders = [folder for folder in folders if folder != most_accurate_folder]
            for folder in other_folders:
                # Get the folder path from the dictionary
                folder_path = structure_dict[folder]

                # Filter the dataframe based on the available options in 'FLOC'
                filtered_dataframe = dataframe[(dataframe['Structure ID'].isin(available_choices))]
                filtered_dataframe = filtered_dataframe[filtered_dataframe['Structure ID'] != most_accurate_folder]
                if no_dist_issue:
                    _, closest_match, available_options = find_closest_match(folder, folder_path, available_choices,
                                                                             filtered_dataframe, verbose=False,
                                                                             resolve=True)
                else:
                    _, closest_match, available_options = find_closest_match(folder, folder_path, available_choices,
                                                                             filtered_dataframe, verbose=False,
                                                                             resolve=True, no_dist_issue=False)
                if closest_match in new_matches and closest_match != match:
                    if isinstance(new_matches[closest_match], list):
                        new_matches[closest_match].append(folder)
                    else:
                        new_matches[closest_match] = [new_matches[closest_match], folder]
                else:
                    new_matches[closest_match] = folder
                    resolved_duplicates[closest_match] = folder

        matches = new_matches  # Update matches for the next iteration

    if verbose:
        print_to_widget("          New closest nadir matches found:")
        for match, folder in resolved_duplicates.items():
            print_to_widget(f"              - {folder} --> {match}")
        print_to_widget("          Make sure that these are correct before proceeding to your next step.")
    return matches, available_choices


def rename_and_zip_directory(progress_queue):
    team_number = team_number_entry.get()
    date = date_entry.get_date().strftime('%m.%d.%Y')
    directory_path = dir_path.get()

    # Validate the team number format (XXXX-YYYY using a simple regex)
    pattern = r'^\d{4}-\d{4}$'
    if not re.match(pattern, team_number):
        messagebox.showerror("Error", "Team Number must be in the format 'XXXX-YYYY'\n(e.g., '0012-0081').")
        return

    # Parse the first 4 integers from team number to get the pilot ID
    pilot_id = team_number[:4]

    # Delete hidden Mac files
    print_to_widget(f"\nDeleting hidden Mac files...")
    found_mac_files = False
    for dirpath, dirnames, filenames in os.walk(directory_path):
        for filename in filenames:
            if filename.startswith("._"):
                filepath = os.path.join(dirpath, filename)
                os.remove(filepath)
                print_to_widget(f"{filename} deleted")
                found_mac_files = True
    if not found_mac_files:
        print_to_widget("No hidden Mac files found in the folder.")
    else:
        print_to_widget(f"All hidden Mac files have been deleted.")

    new_name = f"{team_number}_{date}"

    new_directory_path = os.path.join(os.path.dirname(directory_path), new_name)

    # Check folder names for accuracy and images for Ns
    issues_dict, ezlist = check_issues(directory_path)
    issues_ignored = False
    if any(issues_dict.values()):
        print_to_widget("\n\nWARNING! POTENTIAL ISSUES FOUND!", color='red')
        for issue_type, issues in issues_dict.items():
            # Check if there are any issues of this type
            if issues:
                print_to_widget(f"\n\n{issue_type}", color='red')
                if "500" in issue_type:
                    print_to_widget(f"(Possible causes: Incorrect imagery in folder, incorrect structure inspected, "
                                    f"incorrect GIS coordinates, etc.)\n")
                    for issue in issues:
                        folder_name, detail1, detail2, detail3, detail4 = issue
                        if detail4 > 1:
                            print_to_widget(f"   - {folder_name} (Distance: {detail1} feet; "
                                            f"Possible match: {detail2}, {detail3} feet); Nadir {detail4}")
                        else:
                            print_to_widget(f"   - {folder_name} (Distance: {detail1} feet; "
                                            f"Possible match: {detail2}, {detail3} feet)")
                    print_to_widget(f"\nVerify the images in the folder and the inspected structure."
                                    f"\nIf GIS coordinates are incorrect, write appropriate notes in GIS.")
                elif "GIS STRUCTURE ID" in issue_type:
                    # Detail is a string, assume it's a message
                    print_to_widget(f"(Possible causes: Incorrect folder name, structure not in scope, etc.)\n")
                    for issue in issues:
                        folder_name, detail1 = issue
                        print_to_widget(f"   - {folder_name} (Possible match: {detail1})")
                    print_to_widget(f"\nVerify the folder names and make sure they are in GIS.")
                elif "FLIGHT DATE MISMATCH" in issue_type:
                    # Detail is a datetime object, format it as a date string
                    print_to_widget(f"(Possible causes: Incorrect \"Flight Date\" input above -- ", newline=False)
                    print_to_widget(f"{date_entry.get()}", color='#FFA500', newline=False)
                    print_to_widget(f", incorrect drone or handheld camera dates, etc.)\n")
                    for issue in issues:
                        folder_name, detail1 = issue
                        print_to_widget(f"   - {folder_name} (Date taken: {detail1})")
                    print_to_widget(f"\nVerify that drone and handheld camera dates are set properly.")
                elif "STRUCTURE NADIR" in issue_type:
                    print_to_widget(f"(Possible causes: Nadir image no \"N\" suffix, multiple nadirs, etc.)\n")
                    for issue in issues:
                        folder_name, detail1 = issue
                        if detail1 == 0:
                            print_to_widget(f"   - {folder_name} has no nadir image.")
                        else:
                            print_to_widget(f"   - {folder_name} has {detail1} nadir images.")
                    print_to_widget(f"\nPlease make sure that there is only one nadir per folder, and that it has "
                                    f"an \"N\" suffix.")
                elif "MULTIPLE STRUCTURES" in issue_type:
                    print_to_widget(f"(Possible causes: Incorrect imagery in folder, GPS metadata issues, etc.)\n")
                    for issue in issues:
                        folder_name, detail1, detail2 = issue
                        print_to_widget(f"   - {folder_name} (Farthest distance: {detail1} feet from nadir; "
                                        f"Image name: {detail2})")
                    print_to_widget(f"\nImages in the folder have questionable proximity to the nadir "
                                    f"(Distribution >150 ft; Transmission >500 ft)."
                                    f"\nEnsure that each folder contains only images from a single structure.")
                elif "NADIR GPS" in issue_type:
                    print_to_widget(f"(Possible causes: GPS metadata issues, etc.)\n")
                    for issue in issues:
                        print_to_widget(f"   - {issue}")
                    print_to_widget(f"\nEnsure that all images in all folders have GPS data.")

        print_to_widget("\n\nPLEASE ADDRESS ALL POTENTIAL ISSUES ABOVE AND RERUN THE SCRIPT.", color='#FFA500')
        print_to_widget("(If any exceptions apply or you have ANY questions, please contact QA for support.)")
        # Show a dialog asking the user whether to continue or not
        response = messagebox.askyesno(
            "Potential Issues Found",
            f"Please address all potential issues and then rerun the script."
            f"\n\nDo you want to continue zipping anyway?"
        )
        if not response:  # If the user chooses 'No'
            print_to_widget("\nZipping process terminated.", color='red')
            return
        else:
            issues_ignored = True
    else:
        print_to_widget("\nNo Issues Found!", color='green')

    # Generate the structure_list.txt inside the newly renamed directory
    generate_txt_file(directory_path, pilot_id)

    print_to_widget("\nChecking for EZ Poles...")

    # Function to handle moving EZ Poles to their respective directories
    def handle_ez_poles(ez_list, source_directory):
        """
           Copies all poles into target_directory/Distribution.
           Copies only EZ poles into target_directory/EZPolesForTrans.
        """
        # Create the target directory if it doesn't exist
        if not os.path.exists(source_directory):
            os.makedirs(source_directory)

        # Create separate folders for EZPolesForTrans and Distribution
        ez_poles_dir = os.path.join(source_directory, "EZPolesForTrans")
        all_poles_dir = os.path.join(source_directory, "Distribution")

        # Make sure subfolders exist
        if not os.path.exists(ez_poles_dir):
            os.makedirs(ez_poles_dir)
        if not os.path.exists(all_poles_dir):
            os.makedirs(all_poles_dir)

        innermost_folders = find_innermost_folders(directory_path)

        # Copy only EZ Poles into EZPolesForTrans
        print_to_widget("\nCopying EZ poles into 'EZPolesForTrans' folder...")
        for ez_name in ez_list:
            print(ez_name)
            # Find the corresponding folder path in innermost_folders
            source_folder = next((f for f in innermost_folders if os.path.basename(f) == ez_name), None)
            if source_folder:
                destination_folder = os.path.join(ez_poles_dir, ez_name)
                if source_folder != destination_folder:
                    if os.path.exists(source_folder):
                        print_to_widget(f"  - Copying {ez_name}")
                        try:
                            shutil.copytree(source_folder, destination_folder)
                        except Exception as error:
                            print_to_widget(f"\nError: {error}\n\n"
                                            f"Please close any open files in the directory and try again.", color='red')
                            messagebox.showerror("Error", f"{error}\n\n"
                                                          f"Please close any open files in the directory and try again.")
                    else:
                        print_to_widget(f"Source folder does not exist: {source_folder}")
                else:
                    print_to_widget(f"  - {ez_name} is already in the folder.")

        # Copy *all* poles into Distribution folder
        print_to_widget("\nMoving all poles into 'Distribution' folder...")
        for folder_name in os.listdir(source_directory):
            source_folder = os.path.join(source_directory, folder_name)
            destination_folder = os.path.join(all_poles_dir, folder_name)

            if os.path.isdir(source_folder):
                if folder_name == "EZPolesForTrans" or folder_name == "Distribution":
                    continue

                try:
                    print_to_widget(f"  - Moving {folder_name}")
                    shutil.move(source_folder, destination_folder)
                except Exception as error:
                    print_to_widget(f"\nError moving {folder_name}: {error}\n"
                          f"Please close any open files in the directory and try again.")
                    messagebox.showerror("Error", f"Error moving {folder_name}: {error}")

    if ezlist:
        print_to_widget(f"\nEZ poles found...")
        print(ezlist)
        handle_ez_poles(ezlist, directory_path)
    else:
        print_to_widget(f"\nNo EZ poles found.")

    print_to_widget("\nChecking and deleting empty folders...")
    empty_folders_found = False  # Flag to track if any empty folder is found
    for root_dir, dirs, files in os.walk(directory_path, topdown=False):
        for name in dirs:
            folder_path = os.path.join(root_dir, name)
            try:
                if not os.listdir(folder_path):  # Check if the folder is empty
                    os.rmdir(folder_path)
                    print_to_widget(f"Deleted empty folder: {folder_path}")
                    empty_folders_found = True  # Set the flag to True as an empty folder is found
            except Exception as e:
                print_to_widget(f"\nError: {e}\n\n"
                                f"Please close any open files in the directory and try again.", color='red')
                messagebox.showerror("Error", f"{e}\n\nPlease close any open files in the directory and try again.")
                return
    if not empty_folders_found:
        print_to_widget("No empty folders found.")

    try:
        shutil.move(directory_path, new_directory_path)
    except Exception as e:
        print_to_widget(f"\nError: {e}\n\nPlease close any open files in the directory and try again.", color='red')
        messagebox.showerror("Error", f"{e}\n\nPlease close any open files in the directory and try again.")
        return

    # Calculate the size of the directory
    dir_size = get_directory_size(new_directory_path)
    print_to_widget(f"\nDirectory size: {round(dir_size / 1024 / 1024, 2)} MB.")

    if not issues_ignored:
        print_to_widget("\nData is clean and ready to be zipped. Do you want to proceed?")
        response = messagebox.askyesno("Ready to Zip",
                                       "Data is clean and ready to be zipped.\n\nDo you want to proceed?")
        if not response:
            print_to_widget("Zipping process terminated.", color='red')
            return

    print_to_widget("\nZipping files, please wait...")

    # Name of the ZIP file to create
    zip_name = new_directory_path + '.zip'

    # Start the zipping process in a separate thread
    start_time = time.time()
    zip_thread = threading.Thread(target=zip_directory, args=(new_directory_path, zip_name, progress_queue))
    zip_thread.start()
    cancel_button.configure(state="normal")

    # Continuously update the progress bar until the zipping thread has finished
    while True:
        try:
            progress = progress_queue.get_nowait()
            new_width = (progress / 100) * 448  # Calculate the new width based on the progress percentage
            progress_bar_canvas.coords(progress_bar_rect, 0, 0, new_width, 20)  # Update the width of the progress bar
            progress_bar_canvas.itemconfig(progress_bar_percentage,
                                           text=f"{int(progress)}%")  # Update the percentage print_text
            root.update_idletasks()
        except queue.Empty:
            pass

        if not zip_thread.is_alive() and progress_queue.empty():
            break

        time.sleep(0.1)

    if cancel_zip:
        print_to_widget("\nZipping process cancelled.", color='red')
        messagebox.showinfo("Info", "Zipping process cancelled.")

        # Reset the progress bar to 0%
        progress_bar_canvas.coords(progress_bar_rect, 0, 0, 0, 20)
        progress_bar_canvas.itemconfig(progress_bar_percentage, text="0%")
        time.sleep(2)
        cancel_button.configure(state="disabled")
        os.remove(zip_name)
        return

    # Make sure to process the final progress value (100%)
    try:
        progress = progress_queue.get_nowait()
        progress_bar_canvas.coords(progress_bar_rect, 0, 0, progress * 4.48, 20)
        progress_bar_canvas.itemconfig(progress_bar_percentage, text=f"{int(progress)}%")
        root.update_idletasks()
    except queue.Empty:
        pass

    # Wait for the zipping thread to finish
    zip_thread.join()

    # Stop measuring the time taken to zip the directory
    end_time = time.time()
    total_time = end_time - start_time

    # Get the size of the zipped file in MB
    zip_size = os.path.getsize(zip_name) / 1024 / 1024

    # Calculate the speed of the zipping process in MB/s
    zip_speed = dir_size / total_time / 1024 / 1024

    # Print the total time taken to zip the directory in minutes and seconds
    print_to_widget(f"\nZipping complete!")
    print_to_widget(f"Zipped file size: {round(zip_size, 2)} MB")
    print_to_widget(f"Total time taken: {round(time.time() - start_time, 2)} seconds")
    print_to_widget(f"Zipping speed: {round(zip_speed, 2)} MB/s")

    # Show a message box to inform the user that the data has been packaged successfully
    print_to_widget("\nData has been packaged successfully!", color='green')
    messagebox.showinfo("Success", "Data has been packaged successfully.")

    # Clear the directory entry widget
    dir_path.set('')
    path_entry.delete(0, tk.END)


def choose_directory():
    chosen_directory = filedialog.askdirectory(title="Choose Folder to Package", parent=root)
    dir_path.set(chosen_directory)

    if chosen_directory:
        # List and print folder names in the chosen directory
        print_to_widget("List of Structure IDs in the directory:")
        innermost_folders = []
        for dirpath, dirnames, files in os.walk(chosen_directory):
            # If 'dirnames' is empty, it means 'dirpath' is an innermost folder
            if not dirnames and files:
                folder_name = os.path.basename(dirpath)
                innermost_folders.append(folder_name)
                print_to_widget(folder_name)

        # Reset the progress bar to 0%
        progress_bar_canvas.coords(progress_bar_rect, 0, 0, 0, 20)
        progress_bar_canvas.itemconfig(progress_bar_percentage, text="0%")


closest_distance = None  # Initialize the closest distance variable


def check_issues(directory):
    print_to_widget("\nChecking for potential issues (missing nadir N, incorrect folder names, "
                    "image dates, image coordinates, etc)...")
    # Retrieve the flight date entered
    flight_date = date_entry.get()

    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running in a PyInstaller Bundle
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    excel_file_path = os.path.join(base_path, 'Structure ID List.xlsx')

    # Load the Excel file
    try:
        df = pd.read_excel(excel_file_path)
    except FileNotFoundError:
        try:
            excel_file_path_local = r"F:\2024\Structure ID List.xlsx"
            df = pd.read_excel(excel_file_path_local)
        except FileNotFoundError:
            print_to_widget(f"Error: Structure ID List file not found.", color='red')
            return

    # Remove "OH-" from the beginning of "Structure ID" values
    df['Structure ID'] = df['Structure ID'].astype(str).str.replace('OH-', '', regex=False)

    # Create a set of valid subfolders
    valid_subfolder_set = set(df.iloc[:, 0].dropna())

    # Create a set of EZ Pole structures that are in both scopes
    ez_structures = set(df[df['Structure Type'] == 'EZ_POLE'].iloc[:, 0].dropna())

    # List of image extensions to check
    image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']

    # Initialize a dictionary to store issues
    issues_dict = {
        "FOLDER NAME AND GIS STRUCTURE ID MISMATCH": [],
        "IMAGE METADATA DATE AND FLIGHT DATE MISMATCH": [],
        "UNABLE TO IDENTIFY STRUCTURE NADIR": [],
        "MULTIPLE STRUCTURES IN FOLDER DETECTED": [],
        "STRUCTURE EXCEEDS 500 FEET FROM GIS COORDINATES": [],
        "NADIR GPS DATA NOT FOUND": []
    }
    ez_list = []  # Create lists to store EZ Pole structures

    # Check if directory exists
    if not os.path.exists(directory):
        print_to_widget(f"Error: Directory '{directory}' does not exist.")
        return issues_dict, ez_list

    matches = {}  # Create a dictionary to store the matches
    structure_dict = {}  # Create a dictionary to store the structure names and their paths
    # Iterate over each subfolder in the directory
    for subdir, dirs, files in os.walk(directory):
        if subdir == directory:
            continue
        folder_name = os.path.basename(subdir)
        print_to_widget(f"\nChecking folder {folder_name}...")
        n_coords = None
        structure_dict[folder_name] = subdir

        if not dirs and files:
            if folder_name not in valid_subfolder_set:
                print_to_widget(f"   - Structure ", newline=False)
                print_to_widget(f"{folder_name}", newline=False, color='red')
                print_to_widget(f" not found in GIS. Finding closest match...")
                result = find_closest_match(folder_name, subdir, valid_subfolder_set, df)
                if not result:
                    continue  # Skip to the next iteration of the loop
                else:
                    structure_name_matched, closest_match, options_match = result

                # Check if match already exists
                if closest_match in matches:
                    if not isinstance(matches[closest_match], list):
                        matches[closest_match] = [matches[closest_match]]
                    matches[closest_match].append(structure_name_matched)
                else:
                    matches[closest_match] = structure_name_matched

                # Create a temporary duplicates dictionary for printing purposes
                duplicates = {key: value for key, value in matches.items() if isinstance(value, list)}

                if duplicates:
                    output = "   - The same matches found for the following folders:"
                    for key, values in duplicates.items():
                        output += f"\n              - {', '.join(values)} --> {key}"

                    print_to_widget(f"{output}\n          Finding new matches...")

                    # Resolve duplicates
                    matches, available_options = resolve_duplicates(matches, options_match, structure_dict, df)

                # Find the key in matches that corresponds to the value of structure_name_matched
                matched_key = None
                for key, value in matches.items():
                    if value == folder_name or (isinstance(value, list) and folder_name in value):
                        matched_key = key
                        break

                if matched_key is not None:
                    # Append the tuple (matched_key, folder_name) to the issues_dict
                    issues_dict["FOLDER NAME AND GIS STRUCTURE ID MISMATCH"].append((folder_name, matched_key))
                else:  # if matched_key is None
                    # Append the folder_name to the issues_dict
                    issues_dict["FOLDER NAME AND GIS STRUCTURE ID MISMATCH"].append((folder_name, None))
            else:
                print_to_widget(f"   - ", newline=False)
                print_to_widget(f"Structure {folder_name} found in GIS.")

            # Check where the folder is listed under
            if folder_name in ez_structures:
                ez_list.append(folder_name)

            image_paths = []
            date_taken_list = []
            nadir_count = 0
            valid_subfolder_set_copy = valid_subfolder_set.copy()
            for file in files:
                # Check if the file is an image
                if any(file.lower().endswith(ext) for ext in image_extensions):
                    file_path = os.path.join(subdir, file)
                    image_paths.append(file_path)
                    # Get the date taken from the image
                    date_taken = get_date_taken(file_path)
                    if date_taken and date_taken not in date_taken_list:
                        date_taken_list.append(date_taken)

                    # Check if the image has an 'N' then extract GPS coordinates of that image
                    base_name, extension = os.path.splitext(file)
                    if base_name[-1] == 'N':
                        nadir_count += 1
                        n_coords = get_gps_from_image(file_path)

                        if n_coords:
                            # Find the corresponding row in the DataFrame
                            matching_row = df[df['Structure ID'] == folder_name]
                            # Calculate distance between n_coords and df_coords
                            if not matching_row.empty:
                                df_coords = (matching_row['Latitude'].iloc[0], matching_row['Longitude'].iloc[0])
                                distance = distance_calculator(n_coords, df_coords)
                                if distance is not None:
                                    if distance > 500:
                                        print_to_widget(f"   - Distance from GIS coordinates: ", newline=False)
                                        print_to_widget(f"{distance} feet", newline=False, color='red')
                                        result = find_closest_match(folder_name, subdir, valid_subfolder_set_copy, df,
                                                                    verbose=False, no_dist_issue=False)
                                        if nadir_count > 1:
                                            print_to_widget(f". (Nadir {nadir_count})")
                                        else:  # If there is only one nadir
                                            print_to_widget(f".")
                                        # Check if the result is None
                                        if not result:
                                            issues_dict["STRUCTURE EXCEEDS 500 FEET FROM GIS COORDINATES"].append(
                                                (folder_name, distance, None, None, nadir_count))
                                        else:
                                            structure_name_matched, closest_match, options_match = result

                                            # Check if match already exists
                                            if closest_match in matches:
                                                if not isinstance(matches[closest_match], list):
                                                    matches[closest_match] = [matches[closest_match]]
                                                matches[closest_match].append(structure_name_matched)
                                            else:
                                                matches[closest_match] = structure_name_matched

                                            # Create a temporary duplicates dictionary for printing purposes
                                            duplicates = {key: value for key, value in matches.items() if
                                                          isinstance(value, list)}
                                            if duplicates:
                                                # Resolve duplicates
                                                matches, available_options = resolve_duplicates(
                                                    matches, options_match, structure_dict, df, verbose=False)
                                            # Find key in matches that corresponds to the structure_name_matched value
                                            matched_key = None
                                            for key, value in matches.items():
                                                if (value == folder_name or
                                                        (isinstance(value, list) and folder_name in value)):
                                                    matched_key = key
                                                    break
                                            if matched_key is not None:
                                                (issues_dict["STRUCTURE EXCEEDS 500 FEET FROM GIS COORDINATES"].append(
                                                    (folder_name, distance, matched_key, closest_distance, nadir_count))
                                                )
                                    elif distance >= 150:
                                        print_to_widget(f"   - Distance from GIS coordinates: ", newline=False)
                                        print_to_widget(f"{distance} feet", newline=False, color='#FFA500')
                                        if nadir_count == 1:
                                            print_to_widget(f".")
                                        else:
                                            print_to_widget(f". (Nadir {nadir_count})")
                                    elif distance < 150:  # If the distance is less than 150 feet
                                        if nadir_count == 1:
                                            print_to_widget(f"   - Distance from GIS coordinates: {distance} feet.")
                                        else:
                                            print_to_widget(f"   - Distance from GIS coordinates: {distance} feet. "
                                                            f"(Nadir {nadir_count})")
                        else:
                            print_to_widget(f"   - Nadir image does not have GPS data.", color='red')
                            (issues_dict["NADIR GPS DATA NOT FOUND"].append(folder_name))

            # Check if the date taken is within 24 hours of the flight date
            mismatch_found = False
            for date in date_taken_list:
                if date != flight_date:
                    mismatch_found = True
                    print_to_widget(f"   - Date taken: ", newline=False)
                    print_to_widget(f"{date}.", color='red')
                    issues_dict["IMAGE METADATA DATE AND FLIGHT DATE MISMATCH"].append((folder_name, date))

            if not mismatch_found:
                print_to_widget(f"   - Date taken: {date_taken_list[0]}.")

            # Check if there is an image with 'N' in the folder
            if nadir_count == 0 and subdir != directory:
                print_to_widget(f"   - ", newline=False)
                print_to_widget(f"No nadir", newline=False, color='red')
                print_to_widget(f" image found.")
                issues_dict["UNABLE TO IDENTIFY STRUCTURE NADIR"].append((folder_name, nadir_count))
            elif nadir_count > 1 and subdir != directory:
                print_to_widget(f"   - ", newline=False)
                print_to_widget(f"{nadir_count} nadir", newline=False, color='red')
                print_to_widget(f" images found.")
                issues_dict["UNABLE TO IDENTIFY STRUCTURE NADIR"].append((folder_name, nadir_count))
            else:  # If there is exactly one nadir image
                print_to_widget(f"   - ", newline=False)
                print_to_widget(f"One nadir image found.")

            max_img = None
            if n_coords:
                # Check if multiple structures are found in the folder
                max_distance = 0
                for img_path in image_paths:
                    coord = get_gps_from_image(img_path)
                    img_name = os.path.basename(img_path)
                    if coord:
                        dist_feet = distance_calculator(n_coords, coord)
                        if dist_feet > max_distance:
                            max_distance = dist_feet
                            max_img = img_name
                    else:
                        print_to_widget(f"   - ", newline=False)
                        print_to_widget(f"Warning:", newline=False, color='#FFA500')
                        print_to_widget(f" No GPS data found on {img_name} from structure {folder_name}.")
                # Get the value from the entry widget
                team_number = team_number_entry.get()
                first_digit = team_number[0]
                if first_digit == '1':
                    distance_threshold = 500
                else:
                    distance_threshold = 150
                # Use the distance_threshold in your condition
                if max_distance > distance_threshold:
                    print_to_widget(f"   - Farthest image distance: ", newline=False)
                    print_to_widget(f"{max_distance} feet", newline=False, color='red')
                    print_to_widget(f" from nadir.")
                    (issues_dict["MULTIPLE STRUCTURES IN FOLDER DETECTED"].append(
                        (folder_name, max_distance, max_img)))
                else:
                    print_to_widget(f"   - Farthest image distance: ", newline=False)
                    print_to_widget(f"{max_distance} feet from nadir.")

    return issues_dict, ez_list


# Define the color palette for dark mode
dark_bg = '#18191A'
dark_fg = '#B3B3B3'
accent_color = '#242526'
hover_color = '#3A3B3C'
scrollbar_bg = "#1d1d1d"

# Define the color palette for custom tk dark mode
ctk.set_appearance_mode("Dark")  # Set the appearance to dark mode
ctk.set_default_color_theme("dark-blue")  # Set the default color theme

# Create an instance of the standard tk.Tk class
root = ctk.CTk()
root.title("Package Data")
root.geometry("900x900")

# Add a "Choose Directory" button to the GUI
choose_directory_button = ctk.CTkButton(root, text="1. Choose Directory", command=choose_directory)
choose_directory_button.grid(row=0, column=0, padx=(20, 10), pady=10, sticky='ew')
create_tooltip(choose_directory_button, "Press to select the directory of your structure folders")

# Create a StringVar to store the selected directory path
dir_path = StringVar()

# Add an Entry widget to display and edit the directory path
path_entry = ctk.CTkEntry(root, textvariable=dir_path)
path_entry.grid(row=0, column=1, padx=(0, 20), pady=20, sticky=tk.E + tk.W, columnspan=4)

# Call the update_directory_path function whenever the path_entry widget loses focus or the Enter key is pressed
path_entry.bind('<FocusOut>', lambda _: dir_path.set(path_entry.get().replace('"', '')))
path_entry.bind('<Return>', lambda _: dir_path.set(path_entry.get().replace('"', '')))

# Add the Team Number entry field to the GUI
ctk.CTkLabel(root, text="2. Team Number:").grid(row=1, column=0, padx=5, pady=10, sticky=tk.E)
team_number_entry = ctk.CTkEntry(root, width=80)
team_number_entry.grid(row=1, column=1, columnspan=2, pady=10, sticky=tk.W)

# Add the Flight Date entry field to the GUI
ctk.CTkLabel(root, text="3. Flight Date:").grid(row=1, column=2, padx=5, sticky=tk.E)
date_entry = DateEntry(root, date_pattern='mm.dd.yyyy')
date_entry.configure(background='black', foreground='white', selectbackground='light blue')
date_entry.grid(row=1, column=3, columnspan=2, sticky=tk.W)

# Create a Package Data button
package_data_button = ctk.CTkButton(root, text="4. Package Data", command=packaging_thread_function)
package_data_button.grid(row=2, column=0, columnspan=5, padx=50, pady=25)
create_tooltip(package_data_button, "Press to start packaging the data in the selected directory")

# Create a Cancel Zipping button
cancel_button = ctk.CTkButton(root, text="Cancel Zipping", command=request_cancel)
cancel_button.grid(row=4, column=0, columnspan=5, padx=50, pady=10)
cancel_button.configure(state="disabled")

# Create the "Transmission" label widget
transmission_label = ctk.CTkLabel(root, text="TRANSMISSION:")
transmission_label.grid(row=6, column=3, padx=5, pady=(0, 10), sticky='e')

# Create the "Distribution" label widget
distribution_label = ctk.CTkLabel(root, text="DISTRIBUTION:")
distribution_label.grid(row=7, column=3, padx=5, pady=(0, 10), sticky='e')

# Create the "Field Upload" trans button widget
field_upload_trans_button = ctk.CTkButton(root, text="Field Upload", command=upload_trans_data,
                                          fg_color="#3c6e71", hover_color="#234143")
field_upload_trans_button.grid(row=6, column=4, padx=(0, 20), pady=(0, 10), sticky='e')
create_tooltip(field_upload_trans_button, "Launch Transmission Field Uploads page")

# Create the "Field Upload" distro button widget
field_upload_distro_button = ctk.CTkButton(root, text="Field Upload", command=upload_distro_data, fg_color="#695E93",
                                           hover_color="#504870")
field_upload_distro_button.grid(row=7, column=4, padx=(0, 20), pady=(0, 10), sticky='e')
create_tooltip(field_upload_distro_button, "Launch Distribution Field Uploads page")

# Create a custom progress bar
progress_bar_canvas = ctk.CTkCanvas(root, width=448, height=20, bg="white", highlightthickness=0)
progress_bar_canvas.grid(row=3, column=0, columnspan=5, padx=20, pady=10)
progress_bar_rect = progress_bar_canvas.create_rectangle(0, 0, 0, 30, fill='light green', width=0)
progress_bar_percentage = progress_bar_canvas.create_text(225, 10, text="0%", font=("Arial", 10, "bold"), fill="black",
                                                          anchor="center")

# Create a label widget for the version number
version_label = ctk.CTkLabel(root, text=f"Version {get_current_version()}", cursor="hand2")
version_label.grid(row=8, column=4, padx=10, sticky=tk.E)
version_label.configure(font=("Arial", 10))
version_label.bind("<Button-1>", open_version_history)


# Function to display the context menu
def show_context_menu(event):
    try:
        context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        context_menu.grab_release()


def copy_text(event=None):
    if text_space.selection_get():
        # Copy the selected text to the clipboard
        root.clipboard_clear()
        root.clipboard_append(text_space.selection_get())


# Create the CTkTextbox widget
text_space = ctk.CTkTextbox(root, wrap=tk.WORD)
text_space.configure(fg_color="gray20", font=('Segoe UI', 15))
text_space.grid(row=5, column=0, columnspan=5, padx=20, pady=20, sticky=tk.N + tk.S + tk.E + tk.W)

# Configure the row to expand and fill
root.grid_rowconfigure(5, weight=1)

# Create a context menu
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy", command=copy_text)

# Bind right-click event
text_space.bind("<Button-3>", show_context_menu)


def copy_structure_ids():
    directory = dir_path.get()
    if not directory:
        messagebox.showerror("Error", "No directory selected.")
        return

    innermost_folders = []
    for dirpath, dirnames, _ in os.walk(directory):
        # If 'dirnames' is empty, it means 'dirpath' is an innermost folder
        if not dirnames:
            innermost_folders.append(os.path.basename(dirpath))

    # Join folder names into a string to copy to clipboard
    folder_names_str = '\n'.join(innermost_folders)

    # Copy to clipboard
    pyperclip.copy(folder_names_str)
    messagebox.showinfo("Success", "Structure IDs copied to clipboard. Paste them into the Upload Check sheet.")


# Create the "Copy Structure IDs" button widget
copy_button = ctk.CTkButton(root, text="Copy Structure IDs", width=8, command=copy_structure_ids, fg_color="#565B5E",
                            hover_color="#3a3a3a")
copy_button.grid(row=6, column=0, padx=20, pady=(0, 10), sticky='w')
create_tooltip(copy_button, "Copy structure IDs to paste in Upload Check page")

# Configure the column and row widths to evenly space the buttons
for i in range(3):
    root.grid_columnconfigure(i, weight=1)

# Display the window
root.deiconify()

# Check for update
check_for_updates()

root.mainloop()