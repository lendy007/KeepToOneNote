# Migration of Google Keep notes to Microsoft OneNote - Powershell script

### Thanks to [http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/](http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/) where it has been described how to create OneNote from Powershell

This script can be used for migration of your Google Keep notes into Microsoft OneNote. No guarantees, I have used this script for my own migration and it worked fine for me, so I just decided to share it as I didn't find any useful free tool to performing it.

### Prerequisities

- This script requires Microsoft OneNote application to be installed on computer as it is using COM object of OneNote to create notes.
- You need to visit [takeout.google.com](https://takeout.google.com) and have your Google Keep data exported from Google
- After downloading those exported data you need to extract them and point script to folder where all .html files with notes are located

### Script running
After downloading of script just modify these 2 lines to specify where OneNote notebook with notes will be created. Put just folder there. And you need specify **sourcePath** - path where .html files with exported noted are extracted - usually it should ends as \Takeout\Keep. Then just run the script and wait for results.
```sh
$NotebookPath = ""
$sourcePath = "" 
```
