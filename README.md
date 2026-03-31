# Sonarr-season-folder-normalizer
⚠️ **Note**
This was built entirely with ChatGPT. I have zero coding experience—I just had a problem I needed solved.
Feel free to fork this and improve it.

**Overview:**

This tool detects duplicate season folders (e.g. Season 1 vs Season 01) and inconsistent naming conventions. It provides a preview of issues before optionally fixing them by:

Merging duplicate folders
Renaming folders to a consistent padded format (Season XX)
Screenshot
<img width="2321" height="1207" alt="image" src="https://github.com/user-attachments/assets/2da352b0-aeab-40c2-a575-424426f09aa4" />
Features
Modern GUI with inline confirmations
Recursive scanning of media directories
Detailed output logs
Summary metrics after operations
Optional Sonarr (v4) integration:
Automatically triggers:
RefreshSeries
RescanSeries
Supports path mapping (local ↔ Docker)
Requirements
Python (latest version recommended)
colorama
 (optional – improves GUI appearance)
 
Creates a .json file to store your sonarr api key and url. also creates logging files to your downloads folder. 

**Download:**

season_folder_gui_v4.py

(Optional) Install dependencies

pip install colorama

**Usage**

Run the script:

python season_folder_gui_v4.py
In the GUI:
Browse to your media location
(Optional) Enter your Sonarr IP and API key

**Actions:**

_Scan_

Detects inconsistent season folder naming (e.g. Season 0 vs Season 00)

_Fix_

Renames folders to Season XX format

Merges duplicate folders where necessary
