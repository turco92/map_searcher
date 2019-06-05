# map_searcher
Map ID Searcher for Seven Kingdoms AA (7kfans.com)


7KAA map searcher Readme v1.0

June 1st 2019

Copyright piranhaboy (aka turco_92)

7KAA map searcher v1.0

---
## Table of Contents

I.	What the script do

II.	Requirements

III.	Running the script

IIIa.	Running with .EXE File

IIIb.	Running with .ahk scriptfile

---
### I. What the script do


This script is written with AutoHotkey (v1.1.30.3).
Mainly it opens 7kaa.exe
- moves the window to the top left corner 
- starts a new single player game by clicking with simulated mouse click
- makes a screenshot of the map with the correct map id 
- and starts over

There are two modes of the script:
* Range
* Random
In range mode you enter a starting map id (RangeBegin) and an end id (RangeEnd). The script now starts making screenshots of all maps of your given range.

In random mode you enter a total number of maps you want to search and save. The script uses a random number for every map.

Your screenshots of the maps are saved as .png files in \Desktop\saved_mapID.

---
### II. Requirements


- Windows 10
- Seven Kingdoms: Ancient Adversaries (at least v2.15.1)
- nircmd.exe (for making screenshots, https://www.nirsoft.net/utils/nircmd.html)
- AutoHotkey v1.1.30.03 (ONLY if you want to run .ahk scriptfile)
- patience (if you want save more than 1.000 maps)

---
### III. Running the script


There are two ways to run the script:
- by running .exe file
- by executing .ahk script file

I offer both methods, as nearly all anti-virus programs block the .exe file. Also nircmd.exe is recognized as malicious (I do not now why). Just make an exception in your anti-virus program.

OR

if you do not trust, run the .ahk scriptfile.

You can stop and exit the script any time you want by pressing [CTRL]+[S].

The .ahk scriptfile is actually the source code. It is freeware. You can edit the code as you desire.

---
### IIIa. Running with .EXE File


Just copy map_searcher_v1.x.exe and nircmd.exe to your 7KAA folder and run map_searcher_v1.x.exe

---
### IIIb. Running with .ahk scriptfile


Download and install AutoHotkey: https://www.autohotkey.com/

Copy map_searcher_v1.x.ahk and nircmd.exe to your 7KAA folder. Right-Click on map_searcher_v1.x.ahk > "Run Script"
