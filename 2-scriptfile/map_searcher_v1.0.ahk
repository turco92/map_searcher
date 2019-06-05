;#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
#NoTrayIcon ; Hides Tray icon
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Run script permanent
;#Persistent

; Check windows version
objWMIService := ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\" A_ComputerName "\root\cimv2")
For objOperatingSystem in objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
   Caption := objOperatingSystem.Caption , OSArchitecture := objOperatingSystem.OSArchitecture, CSDVersion  := objOperatingSystem.CSDVersion, Version := objOperatingSystem.Version
;Msgbox %  Caption "`n" OSArchitecture  "`n"  CSDVersion "`n"  Version

if (Caption == "Microsoft Windows 10 Home" or Caption == "Microsoft Windows 10 Pro" or Caption == "Microsoft Windows 10 Education" or Caption == "Microsoft Windows 10 Enterprise") {
	;Win10
} else {
	MsgBox, 48, %Caption%, This script is only tested and confirmed working on Windows 10.
	;ExitApp
}

; Check whether 7kaa.exe is in current directory
if !FileExist("7kaa.exe"){
	MsgBox, 48, Wrong Directory, 7kaa.exe not found in current directory. `nPlease start this script in 7kaa directory!
	ExitApp
}

; Check whether nircmd.exe is in current directory
if !FileExist("nircmd.exe"){
	; Needed for "Help" Button
	Gui, +OwnDialogs
	OnMessage(0x53, "WM_HELP")
	WM_HELP()
	{
		Run, https://www.nirsoft.net/utils/nircmd.html
	}
	
	; MessageBox with "Help"-Button
	message = 
	(
		nircmd.exe not found in current directory. Please check whether nircmd.exe is in the same directory! 
		`nnircmd.exe is used for taking screenshots of maps. Click "help" for more information and download.
	)
	MsgBox, 16432, nircmd.exe Not Found, %message%
	ExitApp
}

; GUI
TextIntro = 
(
	1 - Choose between range or random mode.
	`n2 - Enter Range Begin / Range End OR number of random maps
	`nYou can stop script by pressing [CTRL]+[S] keys. Do not make any mouse clicks during running script! Screenshots are saved in \Desktop\saved_mapID.
)

Gui, Add, Text, x10 y10 w450 h80 , %TextIntro%

; Range or random mode
Gui, Add, Radio, gRange Checked x12 y100 w170 h30, Range
Gui, Add, Radio, gRandom x230 y100 w170 h30 , Random
Gui, Add, Text, vRangeBeginText x12 y140 w80 h30 , RangeBegin
Gui, Add, Text, vRangeEndText x12 y180 w80 h30 , RangeEnd
Gui, Add, Edit, vRangeBegin x90 y140
Gui, Add, Edit, vRangeEnd x90 y180
Gui, Add, Text, vRandomText x230 y140 w120 h30 , Number of random maps
Gui, Add, Edit, vRandomNo x360 y140

;Buttons
Gui, Add, Button, vRangeOK x100 y230 w100 h30 gSubRangeOK, Start
Gui, Add, Button, vRandomOK x100 y230 w100 h30 gSubRandomOK, Start
Gui, Add, Button, x280 y230 w100 h30, Close

; Default: show range mode
GuiControl Hide, RandomText
GuiControl Hide, RandomNo
GuiControl Hide, RandomOK

; Generated using SmartGUI Creator for SciTE
Gui, Show, w500 h280, 7KAA map searcher v1.0
return

GuiClose:
ExitApp

ButtonClose:
ExitApp

Range:
{
    GuiControl Hide, RandomText
    GuiControl Hide, RandomNo
	GuiControl Hide, RandomOK
    GuiControl Show, RangeBegin 
    GuiControl Show, RangeEnd
    GuiControl Show, RangeBeginText
    GuiControl Show, RangeEndText
	GuiControl Show, RangeOK
	return
}

Random:
{
    GuiControl Hide, RangeBeginText
    GuiControl Hide, RangeEndText
    GuiControl Hide, RangeBegin
    GuiControl Hide, RangeEnd
	GuiControl Hide, RangeOK
    GuiControl Show, RandomText
    GuiControl Show, RandomNo
	GuiControl Show, RandomOK
	return
}

SubRangeOK:
{
	Gui, Submit  ; Save the input from the user to each control's associated variable.

	; Setting Mode
	Mode := "Range"
	
	; Ignore comma
	RangeBegin := Floor(RangeBegin)
	RangeEnd := Floor(RangeEnd)

	; Check inputs
	if (RangeBegin <= 0){
		MsgBox, 16, RangeBegin, Please enter a RangeBegin greater than 0!
		Reload
		Exit
	}
	if (RangeEnd <= 0){
		MsgBox, 16, RangeEnd, Please enter a RangeEnd greater than 0!
		Reload
		Exit
	}
	if (RangeBegin > RangeEnd){
		MsgBox, 16, RangeEnd and RangeBegin, RangeEnd must be higher than RangeBegin!
		Reload
		Exit
	}
	if (RangeEnd > 2147483647){
		MsgBox, 16, RangeEnd, RangeEnd can not be greater than 2147483647!
		Reload
		Exit
	}
	; Ask to continue if more than 100 maps
	RangeTotal := RangeEnd - RangeBegin + 1
	if (RangeTotal >= 100){
		MsgBox, 36, Continue?,
		(
			You entered a high range of %RangeTotal% maps.
			`nAre you sure you want to start script?
		)
		IfMsgBox, No
		{
			Reload
			Exit
		}
	}
	
	; Continue to Main
	Goto, Main
}

SubRandomOK:
{
	Gui, Submit  ; Save the input from the user to each control's associated variable.
	
	; Setting Mode
	Mode := "Random"
	
	; Ignore comma
	RandomNo := Floor(RandomNo)
	
	; Check inputs
	if (RandomNo <= 0){
		MsgBox, 16, Number of maps, Please enter a number greater than 0!
		Reload
		Exit
	}
	
	if (RandomNo >= 100){
		MsgBox, 36, Continue?,
		(
			You entered a high number of %RandomNo% maps.
			`nAre you sure you want to start script?
		)
		IfMsgBox, No
		{
			Reload
			Exit
		}
	}
	
	; Continue to Main
	Goto, Main
}

Main:
{
	; Check whether 7K is running and close all processes
	counter := 0
	if WinExist("Seven Kingdoms"){
		WinGet, counter, List, Seven Kingdoms
		while (counter > 0){
			WinKill, Seven Kingdoms
			--counter
		}
	}

	; Mouse blocking
	BlockInput, MouseMove

	; Run Seven Kingdoms
	Run, 7kaa.exe
	Sleep, 1000

	; Check whether 7K is running
	if !WinExist("Seven Kingdoms"){
		MsgBox, 48, 7KAA not running, Seven Kingdoms is not running!
		ExitApp
	}

	; [ALT] + [ENTER] to exit fullscreen mode
	SendInput, !{Enter}

	; Set time to wait between click commands in ms
	waitTime := 400

	; Create saved_mapID folder
	RunWait, %comspec% /c mkdir "%USERPROFILE%\Desktop\saved_mapID", , hide
	
	; Check which Mode is applied
	if(Mode == "Range"){
		counter := 0
	}
	if(Mode == "Random"){
		counter := 1
		RangeTotal := RandomNo
	}
	
	; Set OS
	OS := "Win10"
	
	; Get width and height of window
	width_7K := 0
	height_7K := 0
	WinGetPos, ,,width_7K, height_7K, Seven Kingdoms
	
	; Set x and y border of window in px
	if (OS == "Win10"){
		x_border := 1 + 2 ; AHK is getting wrong border in win10, always 2px more -> +2
		y_border := 26	; title 25 px + border 1 px ; AHK is getting wrong border in win10, always 1px more -> but w/o +1 it`s working fine
		
		; Calculate width w/o border
		width_7k_raw := width_7k - 2 * x_border
		
	} else if (OS == "Win8"){
		x_border := 1
		y_border := 26	; title 25 px + border 1 px
		
		; Calculate width w/o border
		width_7k_raw := width_7k - 2 * x_border
		
	} else if (OS == "Win7"){
		x_border := 3
		y_border := 25	; title 22 px + border 3 px
		
		; Calculate width w/o border
		width_7k_raw := width_7k - 2 * x_border
		
	} else if (OS == "WinXP"){
		x_border := 3
		y_border := 29	; title 26 px + border 3 px
		
		; Calculate width w/o border
		width_7k_raw := width_7k - 2 * x_border
	}
	
	; Set start postion for screenshot 
	if (width_7k_raw == 1024){
		x_start := 753 + x_border
		y_start := 72 + y_border
		x_size := 256
		y_size := 256
	} else if (width_7k_raw == 800){
		x_start := 588 + x_border
		y_start := 56 + y_border
		x_size := 200
		y_size := 200
	}

	; main while-loop
	while (counter <= RangeTotal)
	{
		; Set random map ID
		if(Mode == "Random"){
			Random, RangeBegin, 1, 2147483647
			RangeBegin := Floor(RangeBegin)
		}
		
		; Preparation
		WinActivate, Seven Kingdoms
		Winset, Alwaysontop, On, Seven Kingdoms
		WinMove, Seven Kingdoms, , 0, 0

		; SINGLE PLAYER
		Sleep, %waitTime%
		x := Round((520 / 1030) * width_7K)
		y := Round((320 / 797) * height_7K)
		MouseClick, Left, x, y

		; NEW GAME
		Sleep, %waitTime%
		x := Round((530 / 1030) * width_7K)
		y := Round((390 / 797) * height_7K)
		MouseClick, Left, x, y

		; Advanced Options I
		Sleep, %waitTime%
		x := Round((380 / 1030) * width_7K)
		y := Round((220 / 797) * height_7K)
		MouseClick, Left, x, y

		; Map I.D.
		Sleep, %waitTime%
		Send, %RangeBegin%

		; START
		Sleep, %waitTime%
		x := Round((360 / 1030) * width_7K)
		y := Round((740 / 797) * height_7K)
		MouseClick, Left, x, y

		; Pause game
		Sleep, %waitTime%
		Send, 0

		; Make mapID variable better readable, i.e. 897654367 -> 897-654-367
		if (RangeBegin >= 1000000000){
			var1 := Floor(RangeBegin / 1000000000)
			var2 := Floor((RangeBegin - (var1 * 1000000000)) / 1000000)
			var3 := Floor((RangeBegin - (var1 * 1000000000 + var2 * 1000000)) / 1000)
			var4 := RangeBegin - (var1 * 1000000000 + var2 * 1000000 + var3 * 1000)
			if (var2 < 100){
				var2 := 0 . var2
			}
			if (var3 < 100){
				var3 := 0 . var3
			}
			if (var4 < 100){
				var4 := 0 . var4
			}
			if (var2 < 10){
				var2 := 0 . var2
			}
			if (var3 < 10){
				var3 := 0 . var3
			}
			if (var4 < 10){
				var4 := 0 . var4
			}
			mapID := var1 . "-"
			mapID := mapID . var2
			mapID := mapID . "-"
			mapID := mapID . var3
			mapID := mapID . "-"
			mapID := mapID . var4
		} else if (RangeBegin >= 1000000){
			var1 := Floor(RangeBegin / 1000000)
			var2 := Floor((RangeBegin - (var1 * 1000000)) / 1000)
			var3 := RangeBegin - ((var1 * 1000000) + (var2 * 1000))
			if (var2 < 100){
				var2 := 0 . var2
			}
			if (var3 < 100){
				var3 := 0 . var3
			}
			if (var2 < 10){
				var2 := 0 . var2
			}
			if (var3 < 10){
				var3 := 0 . var3
			}
			mapID := var1 . "-"
			mapID := mapID . var2
			mapID := mapID . "-"
			mapID := mapID . var3
		} else if (RangeBegin >= 1000){
			var1 := Floor(RangeBegin / 1000)
			var2 := RangeBegin - (var1 * 1000)
			if (var2 < 100){
				var2 := 0 . var2
			}
			if (var2 < 10){
				var2 := 0 . var2
			}
			mapID := var1 . "-"
			mapID := mapID . var2
		} else {
			var1 := RangeBegin
			if (RangeBegin < 100){
				var1 := 0 . var1
			}
			if (RangeBegin < 10){
				var1 := 0 . var1
			}
			mapID := var1
		}
				
		; Screenshot of map
		Sleep, %waitTime%
		RunWait, %comspec% /c nircmd savescreenshot "%USERPROFILE%\Desktop\saved_mapID\""%mapID%.png" %x_start% %y_start% %x_size% %y_size%, , hide

		; MENU
		Sleep, %waitTime%
		x := Round((965 / 1030) * width_7K)
		y := Round((60 / 797) * height_7K)
		MouseClick, Left, x, y

		; Quit to Main Menu
		Sleep, %waitTime%
		x := Round((370 / 1030) * width_7K)
		y := Round((545 / 797) * height_7K)
		MouseClick, Left, x, y

		; Yes
		Sleep, %waitTime%
		x := Round((332 / 1030) * width_7K)
		y := Round((540 / 797) * height_7K)
		MouseClick, Left, x, y

		; Set next map ID
		if(Mode == "Range"){
			RangeBegin++
		}
		counter++
	}

	; De-activate "AlwaysOnTop"
	Sleep, %waitTime%
	Winset, Alwaysontop, Off, Seven Kingdoms
	
	; No Mouse blocking
	BlockInput, MouseMoveOff

	; Clean close (all) 7kaa processes
	counter := 0
	if WinExist("Seven Kingdoms"){
		WinGet, counter, List, Seven Kingdoms
		while (counter > 0){
			WinKill, Seven Kingdoms
			--counter
		}
	}

	; Exit script
	ExitApp , 0
}

; [CTRL] + [S] exits script with errorlevel 1
^s::
{
	; No Mouse blocking
	BlockInput, MouseMoveOff
	
	; De-activate "AlwaysOnTop"
	Winset, Alwaysontop, Off, Seven Kingdoms

	; Clean close (all) 7kaa processes
	counter := 0
	if WinExist("Seven Kingdoms"){
		WinGet, counter, List, Seven Kingdoms
		while (counter > 0){
			WinKill, Seven Kingdoms
			--counter
		}
	}

	; Exit script
	ExitApp , 1
}