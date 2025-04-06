#Requires AutoHotkey v2.0
; Copyright (C) 2025 Regy
; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; any later version.

; ==============================================================================
; GLOBAL VARIABLES
; ==============================================================================
global clicking := false
global clickLocations := []
global currentIndex := 1
global isPickingLocation := false
global editingRow := 0
global configFile := A_ScriptDir "\AutoClickerConfig.ini"
global defaultSaveFolder := A_ScriptDir "\SavedConfigs"
global loopStartIndex := 0
global loopEndIndex := 0
global loopCount := 2
global currentLoopIteration := 0

global pickHotkey := "F6"
global startHotkey := "F7"
global prevPickHotkey := pickHotkey
global prevStartHotkey := startHotkey
global hotkeysEnabled := true

; Global flag to track if we're in the process of closing
global isClosing := false
global saveCompleted := false

; Execution amount variables
global executeIndefinitely := false
global executeAmount := 1
global currentExecuteCount := 0
; ==============================================================================
; GUI CREATION
; ==============================================================================
MainWindow := Gui()
MainWindow.Opt("-Resize -MaximizeBox")
MainWindow.Title := "Regy's AC"
MainWindow.MarginX := 15
MainWindow.MarginY := 15

; Create menus
FileMenu := Menu()
FileMenu.Add("Save Click List...", SaveConfig)
FileMenu.Add("Load Click List...", LoadConfig)
FileMenu.Add("Exit", (*) => ExitApp())

SettingsMenu := Menu()
SettingsMenu.Add("Clear All Locations", ClearAllLocations)
SettingsMenu.Add("Configure Hotkeys", OpenHotkeySettings)

LoopMenu := Menu()
LoopMenu.Add("Add Loop Start", AddLoopStartMenu)
LoopMenu.Add("Add Loop End", AddLoopEnd)

MainMenuBar := MenuBar()
MainMenuBar.Add("File", FileMenu)
MainMenuBar.Add("Settings", SettingsMenu)
MainMenuBar.Add("Loop", LoopMenu)
MainWindow.MenuBar := MainMenuBar

; Create main ListView for click locations with better column names
MainWindow.SetFont("s10")
MainWindow.Add("Text", "x15 y15 w300", "Actions List:")
ClickTable := MainWindow.Add("ListView", "x15 y35 w520 h220 Grid -Multi", ["Num", "Type", "Action/Position", "Value", "Target ID"])
ClickTable.ModifyCol(1, 40)
ClickTable.ModifyCol(2, 80)
ClickTable.ModifyCol(3, 100)
ClickTable.ModifyCol(4, 80)
ClickTable.ModifyCol(5, 200)
ClickTable.OnEvent("DoubleClick", EditSelectedLocation)
ClickTable.OnEvent("ContextMenu", ShowContextMenu)

; Execution amount controls
MainWindow.Add("Text", "x15 y260 w80 h20", "Execute:")
ExecuteAmountRadio1 := MainWindow.Add("Radio", "x95 y260 w80 h20 Checked", "Times:")
ExecuteAmountRadio2 := MainWindow.Add("Radio", "x240 y260 w80 h20", "Indefinitely")
ExecuteAmountInput := MainWindow.Add("Edit", "x175 y260 w50 h20", "1")
ExecuteAmountRadio1.OnEvent("Click", UpdateExecutionMode)
ExecuteAmountRadio2.OnEvent("Click", UpdateExecutionMode)

; Add Move Up/Down buttons
MoveUpBtn := MainWindow.Add("Button", "x410 y260 w60 h25", "▲ Up")
MoveUpBtn.OnEvent("Click", MoveItemUp)

MoveDownBtn := MainWindow.Add("Button", "x475 y260 w60 h25", "▼ Down")
MoveDownBtn.OnEvent("Click", MoveItemDown)

; Configuration section
MainWindow.Add("GroupBox", "x15 y282 w520 h90", "Click Configuration")

; Input fields
MainWindow.Add("Text", "x30 y307", "Delay (ms):")
DelayInput := MainWindow.Add("Edit", "x100 y307 w70", "500")

MainWindow.Add("Text", "x200 y307", "Click Type:")
ClickTypeGroup := MainWindow.Add("Radio", "x280 y307 Checked", "Left Click")
MainWindow.Add("Radio", "x280 y332", "Right Click")

; Move the Pick Location button inside the configuration section
PickBtn := MainWindow.Add("Button", "x400 y312 w120 h30", "Pick Location (" pickHotkey ")")
PickBtn.OnEvent("Click", StartPickingLocation)

; Action buttons - Keep start/stop button at bottom
ToggleBtn := MainWindow.Add("Button", "x15 y382 w520 h30", "Start Clicking (" startHotkey ")")
ToggleBtn.OnEvent("Click", ToggleClicking)

; Status bar
StatusBar := MainWindow.Add("StatusBar")
StatusBar.SetText("Right-click on a location to edit or remove it")

; ==============================================================================
; DIALOG WINDOWS
; ==============================================================================
; Click edit dialog - for regular clicks only
ClickEditDialog := Gui("+Owner" MainWindow.Hwnd " +ToolWindow")
ClickEditDialog.Title := "Edit Click Location"
ClickEditDialog.MarginX := 15
ClickEditDialog.MarginY := 10
ClickEditDialog.SetFont("s10")

ClickEditDialog.Add("Text", "x15 y15", "X Position:")
ClickEditXInput := ClickEditDialog.Add("Edit", "x100 y15 w70")

ClickEditDialog.Add("Text", "x15 y45", "Y Position:")
ClickEditYInput := ClickEditDialog.Add("Edit", "x100 y45 w70")

ClickEditDialog.Add("Text", "x15 y75", "Delay (ms):")
ClickEditDelayInput := ClickEditDialog.Add("Edit", "x100 y75 w70")

ClickEditDialog.Add("Text", "x15 y105", "Click Type:")
ClickEditTypeDropDown := ClickEditDialog.Add("DropDownList", "x100 y105 w120", ["Left Click", "Right Click"])

ClickEditDialog.Add("Text", "x15 y135", "Window ID:")
ClickEditWindowIDInput := ClickEditDialog.Add("Edit", "x100 y135 w120 ReadOnly")

SaveClickEditBtn := ClickEditDialog.Add("Button", "x70 y175 w90 h30", "Save")
SaveClickEditBtn.OnEvent("Click", SaveEditedClick)

CancelClickEditBtn := ClickEditDialog.Add("Button", "x170 y175 w90 h30", "Cancel")
CancelClickEditBtn.OnEvent("Click", (*) => ClickEditDialog.Hide())
ClickEditDialog.OnEvent("Close", (*) => ClickEditDialog.Hide())

; Loop edit dialog - for loop start markers only
LoopEditDialog := Gui("+Owner" MainWindow.Hwnd " +ToolWindow")
LoopEditDialog.Title := "Edit Loop"
LoopEditDialog.MarginX := 15
LoopEditDialog.MarginY := 10
LoopEditDialog.SetFont("s10")

LoopEditDialog.Add("Text", "x15 y15", "Loop Count:")
LoopEditCountInput := LoopEditDialog.Add("Edit", "x100 y15 w70")

SaveLoopEditBtn := LoopEditDialog.Add("Button", "x70 y55 w90 h30", "Save")
SaveLoopEditBtn.OnEvent("Click", SaveEditedLoop)

CancelLoopEditBtn := LoopEditDialog.Add("Button", "x170 y55 w90 h30", "Cancel")
CancelLoopEditBtn.OnEvent("Click", (*) => LoopEditDialog.Hide())
LoopEditDialog.OnEvent("Close", (*) => LoopEditDialog.Hide())

; Hotkey dialog
HotkeyDialog := Gui("+Owner" MainWindow.Hwnd " +ToolWindow")
HotkeyDialog.Title := "Configure Hotkeys"
HotkeyDialog.MarginX := 15
HotkeyDialog.MarginY := 10
HotkeyDialog.SetFont("s10")

HotkeyDialog.Add("Text", "x15 y15 w130", "Pick Location:")
PickHotkeyInput := HotkeyDialog.Add("Hotkey", "x150 y15 w120", pickHotkey)

HotkeyDialog.Add("Text", "x15 y45 w130", "Start/Stop Clicking:")
StartHotkeyInput := HotkeyDialog.Add("Hotkey", "x150 y45 w120", startHotkey)

SaveHotkeyBtn := HotkeyDialog.Add("Button", "x70 y85 w100 h30", "Save")
SaveHotkeyBtn.OnEvent("Click", SaveHotkeys)

CancelHotkeyBtn := HotkeyDialog.Add("Button", "x180 y85 w100 h30", "Cancel")
CancelHotkeyBtn.OnEvent("Click", CloseHotkeyDialog)

; Save configuration dialog
SaveConfigDialog := Gui("+Owner" MainWindow.Hwnd " +ToolWindow")
SaveConfigDialog.Title := "Save Click List"
SaveConfigDialog.MarginX := 15
SaveConfigDialog.MarginY := 10
SaveConfigDialog.SetFont("s10")

SaveConfigDialog.Add("Text", "x15 y15", "Click List Name:")
SaveConfigNameInput := SaveConfigDialog.Add("Edit", "x150 y15 w240")

SaveConfigPathDisplay := SaveConfigDialog.Add("Edit", "x15 y45 w300 ReadOnly")
SaveConfigBrowseBtn := SaveConfigDialog.Add("Button", "x325 y45 w80 h25", "Browse...")
SaveConfigBrowseBtn.OnEvent("Click", BrowseSaveLocation)

SaveConfigSaveBtn := SaveConfigDialog.Add("Button", "x120 y80 w100 h30", "Save")
SaveConfigSaveBtn.OnEvent("Click", SaveConfigurationFile)

SaveConfigCancelBtn := SaveConfigDialog.Add("Button", "x230 y80 w100 h30", "Cancel")
SaveConfigCancelBtn.OnEvent("Click", (*) => SaveConfigDialog.Hide())

; Set up dialog event handlers
HotkeyDialog.OnEvent("Close", CloseHotkeyDialog)
SaveConfigDialog.OnEvent("Close", (*) => SaveConfigDialog.Hide())

; ==============================================================================
; UTILITY FUNCTIONS
; ==============================================================================
DisablePickHotkey(disable := false)
{
    global pickHotkey
    
    if (disable) {
        try {
            Hotkey(pickHotkey, "Off")
        } catch as e {
        }
    } else {
        try {
            Hotkey(pickHotkey, "On")
        } catch as e {
        }
    }
}

UpdateListNumbers()
{
    global clickLocations
    
    loopStack := []
    currentMainNumber := 0
    
    Loop ClickTable.GetCount()
    {
        row := A_Index
        location := clickLocations[row]
        displayText := ""
        
        if (location.clickType = "LOOP START") {
            currentMainNumber++
            loopStack.Push(currentMainNumber)
            displayText := currentMainNumber ".0"
        }
        else if (location.clickType = "LOOP END" && loopStack.Length > 0) {
            baseNumber := loopStack[loopStack.Length]
            subNumber := 0
            
            ; Find how many items are inside this loop
            loopStartIndex := 0
            for i, loc in clickLocations {
                if (i < row && loc.clickType = "LOOP START" && loopStartIndex = 0) {
                    loopStartIndex := i
                }
            }
            
            if (loopStartIndex > 0) {
                ; Count non-loop items between start and end
                for i, loc in clickLocations {
                    if (i > loopStartIndex && i < row && loc.clickType != "LOOP START" && loc.clickType != "LOOP END") {
                        subNumber++
                    }
                }
            }
            
            displayText := baseNumber "." (subNumber + 1)
            loopStack.Pop()
        }
        else {
            if (loopStack.Length > 0) {
                baseNumber := loopStack[loopStack.Length]
                
                ; Find the most recent loop start
                loopStartIndex := 0
                for i, loc in clickLocations {
                    if (i < row && loc.clickType = "LOOP START" && i > loopStartIndex) {
                        loopStartIndex := i
                    }
                }
                
                if (loopStartIndex > 0) {
                    ; Count non-loop items between start and this item
                    subNumber := 0
                    for i, loc in clickLocations {
                        if (i > loopStartIndex && i < row && loc.clickType != "LOOP START" && loc.clickType != "LOOP END") {
                            subNumber++
                        }
                    }
                    
                    displayText := baseNumber "." (subNumber + 1)
                }
            }
            else {
                currentMainNumber++
                displayText := currentMainNumber
            }
        }
        
        ; Modify the row directly instead of deleting and inserting
        ClickTable.Modify(row, "Col1", displayText)
    }
}

RefreshClickList()
{
    ; Clear the ListView
    ClickTable.Delete()
    
    ; Re-add all items
    for i, location in clickLocations
    {
        if (location.clickType = "LOOP START") {
            ClickTable.Add("", "", "Loop", "Start", location.loopCount, "")
        } 
        else if (location.clickType = "LOOP END") {
            ClickTable.Add("", "", "Loop", "End", "", "")
        }
        else {
            posDisplay := location.x "," location.y
            ClickTable.Add("", "", location.clickType, posDisplay, location.delay, location.winHwnd)
        }
    }
    
    ; Update numbering and save
    UpdateListNumbers()
    SaveClickLocations()
}

UpdateExecutionMode(*)
{
    global executeIndefinitely, ExecuteAmountInput
    
    if (ExecuteAmountRadio1.Value)
    {
        executeIndefinitely := false
        ExecuteAmountInput.Enabled := true
    }
    else
    {
        executeIndefinitely := true
        ExecuteAmountInput.Enabled := false
    }
}

; ==============================================================================
; LOOP FUNCTIONS
; ==============================================================================
; Checks if we are currently inside a loop
IsInsideLoop()
{
    global clickLocations
    
    loopStartCount := 0
    loopEndCount := 0
    
    for i, location in clickLocations
    {
        if (location.clickType = "LOOP START")
            loopStartCount++
        else if (location.clickType = "LOOP END")
            loopEndCount++
    }
    
    ; If we have more starts than ends, we're inside a loop
    return (loopStartCount > loopEndCount)
}

AddLoopStartMenu(*)
{
    global clickLocations, loopCount
    
    ; Check if we're already inside a loop
    if (IsInsideLoop())
    {
        MsgBox("Cannot create nested loops. Please close the current loop first.", "Loop Error", "Icon!")
        return
    }
    
    loopCountInput := InputBox("Enter the number of loop iterations:", "Loop Count", "w200 h100", "2")
    if (loopCountInput.Result != "OK")
        return
    
    loopCount := loopCountInput.Value
    
    if !IsInteger(loopCount) || loopCount < 1
    {
        MsgBox("Please enter a valid loop count (minimum 1)")
        return
    }
    
    ; Add with new formatting - "Loop" in Type column, "Start" in Action column
    ClickTable.Add("", "", "Loop", "Start", loopCount, "")
    
    clickLocations.Push({
        x: 0, 
        y: 0, 
        delay: loopCount, 
        clickType: "LOOP START",
        winTitle: "Loop begins",
        winHwnd: "",
        loopType: "Start",
        loopCount: loopCount
    })
    
    SaveClickLocations()
    UpdateListNumbers()
    
    StatusBar.SetText("Added loop start with " loopCount " iterations")
}

AddLoopEnd(*)
{
    global clickLocations
    
    ; First check if we have any loop start without a matching end
    loopStartCount := 0
    loopEndCount := 0
    
    for i, location in clickLocations
    {
        if (location.clickType = "LOOP START")
            loopStartCount++
        else if (location.clickType = "LOOP END")
            loopEndCount++
    }
    
    if (loopStartCount <= loopEndCount)
    {
        MsgBox("Cannot add loop end. There is no open loop to close.", "Loop Error", "Icon!")
        return
    }
    
    ; Add with new formatting - "Loop" in Type column, "End" in Action column
    ClickTable.Add("", "", "Loop", "End", "", "")
    
    clickLocations.Push({
        x: 0, 
        y: 0, 
        delay: 100, 
        clickType: "LOOP END",
        winTitle: "Loop ends",
        winHwnd: "",
        loopType: "End"
    })
    
    SaveClickLocations()
    UpdateListNumbers()
    
    StatusBar.SetText("Added loop end marker")
}

; ==============================================================================
; HOTKEY FUNCTIONS
; ==============================================================================
ToggleHotkeys(enable := true)
{
    global hotkeysEnabled, clicking, pickHotkey, startHotkey
    
    if (enable && !hotkeysEnabled)
    {
        try {
            Hotkey(pickHotkey, "On")
        } catch as e {
        }
        
        try {
            Hotkey(startHotkey, "On") 
        } catch as e {
        }
        
        try {
            Hotkey("^Up", "On")
        } catch as e {
        }
        
        try {
            Hotkey("^Down", "On")
        } catch as e {
        }
        
        hotkeysEnabled := true
    }
    else if (!enable && hotkeysEnabled)
    {
        try {
            Hotkey(pickHotkey, "Off")
        } catch as e {
        }
        
        try {
            Hotkey(startHotkey, "Off")
        } catch as e {
        }
        
        try {
            Hotkey("^Up", "Off")
        } catch as e {
        }
        
        try {
            Hotkey("^Down", "Off")
        } catch as e {
        }
        
        hotkeysEnabled := false
    }
}

SetupHotkeys()
{
    global pickHotkey, startHotkey
    global prevPickHotkey, prevStartHotkey
    
    try {
        if (prevPickHotkey != "")
            Hotkey(prevPickHotkey, "Off")
    } catch as e {
    }
    
    try {
        if (prevStartHotkey != "")
            Hotkey(prevStartHotkey, "Off")
    } catch as e {
    }
    
    try {
        Hotkey(pickHotkey, TogglePickingLocation)
        prevPickHotkey := pickHotkey
    } catch as e {
        MsgBox("Error setting pick location hotkey: " e.Message)
    }
    
    try {
        Hotkey(startHotkey, ToggleClicking)
        prevStartHotkey := startHotkey
    } catch as e {
        MsgBox("Error setting start/stop clicking hotkey: " e.Message)
    }
    
    try {
        Hotkey("^Up", MoveItemUp)      ; Ctrl+Up to move item up
        Hotkey("^Down", MoveItemDown)  ; Ctrl+Down to move item down
    } catch as e {
        MsgBox("Error setting move item hotkeys: " e.Message)
    }
    
    hotkeysEnabled := true
}

OpenHotkeySettings(*)
{
    ToggleHotkeys(false)
    
    PickHotkeyInput.Value := pickHotkey
    StartHotkeyInput.Value := startHotkey
    
    StatusBar.SetText("Hotkeys disabled while configuring")
    
    HotkeyDialog.Show()
}

CloseHotkeyDialog(*)
{
    HotkeyDialog.Hide()
    
    ToggleHotkeys(true)
    
    StatusBar.SetText("Right-click on a location to edit or remove it")
}

SaveHotkeys(*)
{
    global pickHotkey, startHotkey
    global prevPickHotkey, prevStartHotkey
    
    newPickHotkey := PickHotkeyInput.Value
    newStartHotkey := StartHotkeyInput.Value
    
    if (newPickHotkey = "" || newStartHotkey = "") {
        MsgBox("All hotkeys must be specified!")
        return
    }
    
    if (newPickHotkey = newStartHotkey) {
        MsgBox("All hotkeys must be different!")
        return
    }
    
    prevPickHotkey := pickHotkey
    prevStartHotkey := startHotkey
    
    pickHotkey := newPickHotkey
    startHotkey := newStartHotkey
    
    PickBtn.Text := "Pick Location (" pickHotkey ")"
    ToggleBtn.Text := clicking ? "Stop Clicking (" startHotkey ")" : "Start Clicking (" startHotkey ")"
    
    SetupHotkeys()

    SaveSettings()
    
    CloseHotkeyDialog()
}

; ==============================================================================
; CONTEXT MENU FUNCTIONS
; ==============================================================================
ShowContextMenu(LV, RowNumber, IsRightClick, X, Y)
{
    if (RowNumber > 0)
    {
        clickType := clickLocations[RowNumber].clickType
        
        if (clickType = "LOOP END")
        {
            ; Special menu for loop ends
            LoopEndContextMenu := Menu()
            if (RowNumber > 1)
                LoopEndContextMenu.Add("Move Up", MoveItemUp)
            if (RowNumber < clickLocations.Length)
                LoopEndContextMenu.Add("Move Down", MoveItemDown)
            LoopEndContextMenu.Add("Remove Loop End", RemoveSelectedLocation)
            LoopEndContextMenu.Show()
        }
        else if (clickType = "LOOP START")
        {
            ; Special menu for loop starts
            LoopStartContextMenu := Menu()
            if (RowNumber > 1)
                LoopStartContextMenu.Add("Move Up", MoveItemUp)
            if (RowNumber < clickLocations.Length)
                LoopStartContextMenu.Add("Move Down", MoveItemDown)
            LoopStartContextMenu.Add("Edit Loop", EditSelectedLocation)
            LoopStartContextMenu.Add("Remove Loop Start", RemoveSelectedLocation)
            LoopStartContextMenu.Show()
        }
        else
        {
            ; Regular clicks menu
            ClickContextMenu := Menu()
            if (RowNumber > 1)
                ClickContextMenu.Add("Move Up", MoveItemUp)
            if (RowNumber < clickLocations.Length)
                ClickContextMenu.Add("Move Down", MoveItemDown)
            ClickContextMenu.Add("Edit Click", EditSelectedLocation)
            ClickContextMenu.Add("Remove Click", RemoveSelectedLocation)
            ClickContextMenu.Show()
        }
    }
}

; ==============================================================================
; LOCATION PICKING FUNCTIONS
; ==============================================================================
TogglePickingLocation(*)
{
    global isPickingLocation, clicking
    
    if (clicking && !isPickingLocation) {
        MsgBox("Please stop clicking before picking a new location")
        return
    }
    
    if (isPickingLocation)
    {
        CancelPickingLocation()
    }
    else
    {
        StartPickingLocation()
    }
}

StartPickingLocation(*)
{
    global isPickingLocation, clicking
    
    if (clicking) {
        MsgBox("Please stop clicking before picking a new location")
        return
    }
    
    isPickingLocation := true
    
    PickBtn.Text := "Cancel Picking (" pickHotkey ")"
    PickBtn.Enabled := true
    
    MainWindow.Hide()
    
    SetTimer(CheckForClick, 10)
}

CancelPickingLocation()
{
    global isPickingLocation
    
    isPickingLocation := false
    
    SetTimer(CheckForClick, 0)
    
    MainWindow.Show()
    
    PickBtn.Text := "Pick Location (" pickHotkey ")"
    PickBtn.Enabled := true
}

CheckForClick()
{
    global isPickingLocation
    
    if (!isPickingLocation)
        return
        
    if (GetKeyState("LButton", "P"))
    {
        MouseGetPos(&xpos, &ypos, &winHwnd, &winControl)
        
        winTitle := WinGetTitle("ahk_id " winHwnd)
        winProcessName := WinGetProcessName("ahk_id " winHwnd)
        
        winIdentifier := "ahk_id " winHwnd
        
        if (winProcessName)
            winDisplayName := winProcessName
        else if (winTitle) {
            winDisplayName := StrLen(winTitle) > 30 ? SubStr(winTitle, 1, 27) "..." : winTitle
        }
        else
            winDisplayName := "Unknown Window"
        
        isPickingLocation := false
        
        SetTimer(CheckForClick, 0)
        
        MainWindow.Show()
        
        PickBtn.Text := "Pick Location (" pickHotkey ")"
        PickBtn.Enabled := true
        
        AddLocation(xpos, ypos, winDisplayName, winHwnd)
    }
}

AddLocation(xpos, ypos, winTitle, winHwnd)
{
    delay := DelayInput.Value
    
    if !IsInteger(delay) || delay < 10
    {
        MsgBox("Please enter a valid delay (at least 10ms)")
        return
    }
    
    if (xpos = 0 && ypos = 0)
        return
    
    clickType := ClickTypeGroup.Value = 1 ? "Left Click" : "Right Click"
    
    posDisplay := xpos "," ypos
    
    ; Add with new column layout - Type, Action/Position, Value, Target ID
    ClickTable.Add("", "", clickType, posDisplay, delay, winHwnd)
    
    clickLocations.Push({
        x: xpos, 
        y: ypos, 
        delay: delay, 
        clickType: clickType,
        winTitle: winTitle,  ; Still stored but not shown in ListView
        winHwnd: winHwnd,
        loopType: ""
    })

    SaveClickLocations()
    UpdateListNumbers()
    
    StatusBar.SetText("Added new click location: " posDisplay)
}

; ==============================================================================
; CLICKING FUNCTIONS
; ==============================================================================
ToggleClicking(*)
{
    global clicking, ToggleBtn
    global loopStartIndex, loopEndIndex, currentLoopIteration, currentIndex
    global executeIndefinitely, executeAmount, currentExecuteCount
    
    if (clickLocations.Length = 0)
    {
        MsgBox("No click locations added!")
        return
    }
    
    if (clicking)
    {
        clicking := false
        SetTimer(PerformClicks, 0)
        ToggleBtn.Text := "Start Clicking (" startHotkey ")"
        StatusBar.SetText("Clicking stopped")
        
        loopStartIndex := 0
        loopEndIndex := 0
        currentLoopIteration := 0
        currentExecuteCount := 0
        
        PickBtn.Enabled := true
        DisablePickHotkey(false)
        
        ; Re-enable execution controls
        ExecuteAmountRadio1.Enabled := true
        ExecuteAmountRadio2.Enabled := true
        if (!executeIndefinitely)
            ExecuteAmountInput.Enabled := true
    }
    else
    {
        ; Read execution amount if not set to indefinite
        if (!executeIndefinitely)
        {
            executeAmount := ExecuteAmountInput.Value
            if (!IsInteger(executeAmount) || executeAmount < 1)
            {
                MsgBox("Please enter a valid execution amount (minimum 1)")
                return
            }
        }
        
        clicking := true
        ToggleBtn.Text := "Stop Clicking (" startHotkey ")"
        
        ; Always reset to the first click position when starting
        currentIndex := 1
        loopStartIndex := 0
        loopEndIndex := 0
        currentLoopIteration := 0
        currentExecuteCount := 0
        
        ; Disable execution controls while running
        ExecuteAmountRadio1.Enabled := false
        ExecuteAmountRadio2.Enabled := false
        ExecuteAmountInput.Enabled := false
        
        if (executeIndefinitely)
            StatusBar.SetText("Clicking started - running indefinitely")
        else
            StatusBar.SetText("Clicking started - will execute " executeAmount " times")
        
        PickBtn.Enabled := false
        DisablePickHotkey(true)
        
        SetTimer(PerformClicks, 10)
    }
}

PerformClicks()
{
    global clicking, clickLocations, currentIndex
    global loopStartIndex, loopEndIndex, loopCount, currentLoopIteration
    global executeIndefinitely, executeAmount, currentExecuteCount
    
    if (!clicking)
        return
    
    location := clickLocations[currentIndex]
    
    ; Check if this is the first element in a new execution cycle
    if (currentIndex == 1 && currentExecuteCount > 0 && !executeIndefinitely)
    {
        ; Check if we've completed all executions
        if (currentExecuteCount >= executeAmount)
        {
; Stop clicking - we've completed the requested number of executions
            clicking := false
            SetTimer(PerformClicks, 0)
            ToggleBtn.Text := "Start Clicking (" startHotkey ")"
            StatusBar.SetText("Clicking completed - executed " executeAmount " times")
            
            loopStartIndex := 0
            loopEndIndex := 0
            currentLoopIteration := 0
            currentExecuteCount := 0
            
            PickBtn.Enabled := true
            DisablePickHotkey(false)
            
            ; Re-enable execution controls
            ExecuteAmountRadio1.Enabled := true
            ExecuteAmountRadio2.Enabled := true
            if (!executeIndefinitely)
                ExecuteAmountInput.Enabled := true
                
            return
        }
    }
    
    if (location.clickType = "LOOP START")
    {
        loopStartIndex := currentIndex
        loopCount := location.loopCount
        currentLoopIteration := 0
        
        currentIndex := Mod(currentIndex, clickLocations.Length) + 1
        SetTimer(PerformClicks, 10)
        return
    }
    else if (location.clickType = "LOOP END")
    {
        if (currentLoopIteration >= loopCount - 1)
        {
            currentIndex := Mod(currentIndex, clickLocations.Length) + 1
            
            ; Check if we just completed a full execution
            if (currentIndex == 1)
            {
                currentExecuteCount++
                if (!executeIndefinitely)
                    StatusBar.SetText("Execution " currentExecuteCount " of " executeAmount)
            }
        }
        else
        {
            currentLoopIteration++
            currentIndex := loopStartIndex + 1
            StatusBar.SetText("Loop iteration " currentLoopIteration + 1 " of " loopCount)
        }
        
        SetTimer(PerformClicks, 10)
        return
    }
    
    if WinExist("ahk_id " location.winHwnd)
    {
        WinActivate("ahk_id " location.winHwnd)
        
        Sleep(50)
        
        BlockInput(true)
        MouseMove(location.x, location.y)
        
        if (location.clickType = "Left Click")
            Click()
        else if (location.clickType = "Right Click")
            Click("Right")
        
        BlockInput(false)
    }
    else
    {
        if (location.winTitle && WinExist(location.winTitle))
        {
            WinActivate(location.winTitle)
            
            Sleep(50)
            
            BlockInput(true)
            MouseMove(location.x, location.y)
            
            if (location.clickType = "Left Click")
                Click()
            else if (location.clickType = "Right Click")
                Click("Right")
            
            BlockInput(false)
        }
        else
        {
            StatusBar.SetText("Window for click #" currentIndex " no longer exists!")
            
            currentIndex := Mod(currentIndex, clickLocations.Length) + 1
            
            ; Check if we just completed a full execution
            if (currentIndex == 1)
            {
                currentExecuteCount++
                if (!executeIndefinitely)
                    StatusBar.SetText("Execution " currentExecuteCount " of " executeAmount)
            }
            
            SetTimer(PerformClicks, 10)
            return
        }
    }
    
    currentIndex := Mod(currentIndex, clickLocations.Length) + 1
    
    ; Check if we just completed a full execution
    if (currentIndex == 1)
    {
        currentExecuteCount++
        if (!executeIndefinitely)
            StatusBar.SetText("Execution " currentExecuteCount " of " executeAmount)
    }
    
    SetTimer(PerformClicks, location.delay)
}

; ==============================================================================
; EDIT AND REMOVE FUNCTIONS
; ==============================================================================
EditSelectedLocation(*)
{
    rowNumber := ClickTable.GetNext(0)
    
    if (rowNumber > 0)
    {
        location := clickLocations[rowNumber]
        
        if (location.clickType = "LOOP END")
        {
            MsgBox("Loop end markers cannot be edited.", "Not Editable", "Icon!")
            return
        }
        else if (location.clickType = "LOOP START")
        {
            OpenLoopEditDialog(rowNumber)
        }
        else
        {
            OpenClickEditDialog(rowNumber)
        }
    }
}

OpenClickEditDialog(rowNumber)
{
    global editingRow
    
    editingRow := rowNumber
    
    location := clickLocations[rowNumber]
    
    ; Set click edit fields
    ClickEditXInput.Value := location.x
    ClickEditYInput.Value := location.y
    ClickEditDelayInput.Value := location.delay
    ClickEditWindowIDInput.Value := location.winHwnd
    
    ; Handle click type selection
    if (location.clickType = "Left Click")
        ClickEditTypeDropDown.Choose(1)
    else if (location.clickType = "Right Click")
        ClickEditTypeDropDown.Choose(2)
    
    ClickEditDialog.Show()
}

OpenLoopEditDialog(rowNumber)
{
    global editingRow
    
    editingRow := rowNumber
    
    location := clickLocations[rowNumber]
    
    ; Set loop count
    if (location.HasProp("loopCount"))
        LoopEditCountInput.Value := location.loopCount
    else
        LoopEditCountInput.Value := "2"
    
    LoopEditDialog.Show()
}

SaveEditedClick(*)
{
    global editingRow
    
    xpos := ClickEditXInput.Value
    ypos := ClickEditYInput.Value
    delay := ClickEditDelayInput.Value
    clickType := ClickEditTypeDropDown.Text
    winHwnd := ClickEditWindowIDInput.Value
    
    ; Validate coordinates and delay
    if !IsInteger(xpos) || !IsInteger(ypos)
    {
        MsgBox("Please enter valid coordinates")
        return
    }
    
    if !IsInteger(delay) || delay < 10
    {
        MsgBox("Please enter a valid delay (at least 10ms)")
        return
    }
    
    ; Get window title from hwnd for display purposes
    winTitle := "Unknown Window"
    if (winHwnd)
    {
        try {
            winTitle := WinGetTitle("ahk_id " winHwnd)
            if (!winTitle)
                winTitle := WinGetProcessName("ahk_id " winHwnd)
        }
        catch {
            winTitle := "Unknown Window"
        }
    }
    
    posDisplay := xpos "," ypos
    
    ; Update ListView - Use the correct column indexes
    ; Column numbers are: 1=Num, 2=Type, 3=Action/Position, 4=Value, 5=Target ID
    ClickTable.Modify(editingRow, "Col2", clickType)
    ClickTable.Modify(editingRow, "Col3", posDisplay)
    ClickTable.Modify(editingRow, "Col4", delay)
    ClickTable.Modify(editingRow, "Col5", winHwnd)
    
    ; Update internal data
    newLocation := {
        x: xpos, 
        y: ypos, 
        delay: delay, 
        clickType: clickType,
        winTitle: winTitle,
        winHwnd: winHwnd,
        loopType: ""
    }
    
    clickLocations[editingRow] := newLocation
    
    ClickEditDialog.Hide()

    SaveClickLocations()
    UpdateListNumbers()
    
    StatusBar.SetText("Click location updated")
}

SaveEditedLoop(*)
{
    global editingRow
    
    loopCount := LoopEditCountInput.Value
    
    ; Validate loop count
    if !IsInteger(loopCount) || loopCount < 1
    {
        MsgBox("Please enter a valid loop count (at least 1)")
        return
    }
    
    ; Update ListView with new format - use the correct column
    ClickTable.Modify(editingRow, "Col4", loopCount)
    
    ; Update internal data
    location := clickLocations[editingRow]
    location.loopCount := loopCount
    location.delay := loopCount
    
    LoopEditDialog.Hide()

    SaveClickLocations()
    UpdateListNumbers()
    
    StatusBar.SetText("Loop count updated to " loopCount)
}

MoveItemUp(*)
{
    global clickLocations
    
    rowNumber := ClickTable.GetNext(0)
    
    if (rowNumber <= 1 || rowNumber > clickLocations.Length)
        return
    
    ; Swap items in the clickLocations array
    temp := clickLocations[rowNumber]
    clickLocations[rowNumber] := clickLocations[rowNumber - 1]
    clickLocations[rowNumber - 1] := temp
    
    ; Update the ListView display - do this by reloading the entire list
    RefreshClickList()
    
    ; Select the moved item at its new position
    ClickTable.Modify(rowNumber - 1, "Select Focus")
}

MoveItemDown(*)
{
    global clickLocations
    
    rowNumber := ClickTable.GetNext(0)
    
    if (rowNumber < 1 || rowNumber >= clickLocations.Length)
        return
    
    ; Swap items in the clickLocations array
    temp := clickLocations[rowNumber]
    clickLocations[rowNumber] := clickLocations[rowNumber + 1]
    clickLocations[rowNumber + 1] := temp
    
    ; Update the ListView display - do this by reloading the entire list
    RefreshClickList()
    
    ; Select the moved item at its new position
    ClickTable.Modify(rowNumber + 1, "Select Focus")
}

RemoveSelectedLocation(*)
{
    rowNumber := ClickTable.GetNext(0)
    
    if (rowNumber > 0)
    {
        RemoveSpecificLocation(rowNumber)
        SaveClickLocations()
    }
}

RemoveSpecificLocation(rowNumber)
{
    global clickLocations, currentIndex
    
    ClickTable.Delete(rowNumber)
    clickLocations.RemoveAt(rowNumber)
    
    if (clickLocations.Length = 0)
    {
        currentIndex := 1
        IniDelete(configFile, "ClickLocations")
        IniWrite(0, configFile, "ClickLocations", "Count")
    }
    else if (currentIndex > clickLocations.Length)
        currentIndex := clickLocations.Length
    else
        SaveClickLocations()
        
    UpdateListNumbers()
}

ClearAllLocations(*)
{
    global clicking, clickLocations, currentIndex, loopStartIndex, loopEndIndex, currentLoopIteration
    
    if (clicking)
    {
        MsgBox("Stop clicking before clearing all locations")
        return
    }
    
    result := MsgBox("Are you sure you want to clear all locations?", "Confirm Clear", "YesNo")
    if (result != "Yes")
        return
        
    ClickTable.Delete()
    
    clickLocations := Array()
    currentIndex := 1
    loopStartIndex := 0
    loopEndIndex := 0
    currentLoopIteration := 0

    IniDelete(configFile, "ClickLocations")
    IniWrite(0, configFile, "ClickLocations", "Count")
    
    StatusBar.SetText("All locations cleared")
    
    UpdateListNumbers()
}

; ==============================================================================
; CONFIGURATION FUNCTIONS
; ==============================================================================
SaveSettings()
{
    global configFile
    
    IniWrite(pickHotkey, configFile, "Hotkeys", "PickLocation")
    IniWrite(startHotkey, configFile, "Hotkeys", "StartStopClicking")

    IniWrite(DelayInput.Value, configFile, "Settings", "DefaultDelay")
    
    IniWrite(ClickTypeGroup.Value, configFile, "Settings", "DefaultClickType")
}

LoadSavedSettings()
{
    global configFile
    global pickHotkey, startHotkey
    global prevPickHotkey, prevStartHotkey
    
    if !FileExist(configFile)
        return
    
    try
    {
        pickHotkey := IniRead(configFile, "Hotkeys", "PickLocation", "F6")
        
        try {
            startHotkey := IniRead(configFile, "Hotkeys", "StartStopClicking", "F7")
        } catch {
            startHotkey := IniRead(configFile, "Hotkeys", "StartClicking", "F7")
        }
        
        prevPickHotkey := pickHotkey
        prevStartHotkey := startHotkey
        
        PickBtn.Text := "Pick Location (" pickHotkey ")"
        ToggleBtn.Text := "Start Clicking (" startHotkey ")"
    }
    catch
    {
    }
    
    try
    {
        defaultDelay := IniRead(configFile, "Settings", "DefaultDelay", "500")
        DelayInput.Value := defaultDelay
    }
    catch
    {
    }
    
    try
    {
        defaultClickType := IniRead(configFile, "Settings", "DefaultClickType", "1")
        ClickTypeGroup.Value := defaultClickType
    }
    catch
    {
    }
}

SaveClickLocations()
{
    global configFile
    
    IniDelete(configFile, "ClickLocations")
    
    IniWrite(clickLocations.Length, configFile, "ClickLocations", "Count")
    
    if (clickLocations.Length = 0)
        return
    
    for i, location in clickLocations
    {
        loopInfo := location.HasProp("loopType") ? location.loopType : ""
        loopCount := (location.HasProp("loopCount") && location.loopCount) ? location.loopCount : "0"
        
        locationStr := location.x "," location.y "," location.delay "," location.clickType 
                    . "," location.winTitle "," location.winHwnd "," loopInfo "," loopCount
                    
        IniWrite(locationStr, configFile, "ClickLocations", "Location" i)
    }
}

; ==============================================================================
; FILE OPERATIONS
; ==============================================================================
SaveConfig(*)
{
    global defaultSaveFolder
    
    if (clickLocations.Length = 0)
    {
        MsgBox("No click locations to save!")
        return
    }
    
    SaveConfigPathDisplay.Value := defaultSaveFolder
    SaveConfigNameInput.Value := "MyClickList"
    
    SaveConfigDialog.Show()
}

BrowseSaveLocation(*)
{
    selectedDir := DirSelect("*" A_ScriptDir, 3, "Select folder to save configuration")
    
    if (selectedDir != "")
        SaveConfigPathDisplay.Value := selectedDir
}

SaveConfigurationFile(*)
{
    savePath := SaveConfigPathDisplay.Value
    configName := SaveConfigNameInput.Value
    
    if (configName = "")
    {
        MsgBox("Please enter a configuration name")
        return
    }
    
    configName := RegExReplace(configName, "[\\/:*?`"<>|]", "_")
    
    ; Changed file extension from .click to .rcl
    filePath := savePath "\" configName ".rcl"
    
    try
    {
        FileObj := FileOpen(filePath, "w")
        
        FileObj.WriteLine("[AutoClickerList]")
        FileObj.WriteLine("Version=1.1")
        FileObj.WriteLine("SaveDate=" A_Now)
        FileObj.WriteLine("LocationCount=" clickLocations.Length)
        
        for i, location in clickLocations
        {
            FileObj.WriteLine("[Location" i "]")
            FileObj.WriteLine("X=" location.x)
            FileObj.WriteLine("Y=" location.y)
            FileObj.WriteLine("Delay=" location.delay)
            FileObj.WriteLine("ClickType=" location.clickType)
            
            if (location.HasProp("loopType") && location.loopType)
                FileObj.WriteLine("LoopType=" location.loopType)
                
            if (location.HasProp("loopCount") && location.loopCount)
                FileObj.WriteLine("LoopCount=" location.loopCount)
                
            FileObj.WriteLine("WindowTitle=" location.winTitle)
            FileObj.WriteLine("WindowID=" location.winHwnd)
        }
        
        FileObj.Close()
        
        MsgBox("Click list saved to " filePath)
        SaveConfigDialog.Hide()
        StatusBar.SetText("Click list saved to: " filePath)
    }
    catch as e
    {
        MsgBox("Error saving configuration: " e.Message)
        StatusBar.SetText("Error saving configuration")
    }
}

LoadConfig(*)
{
    global defaultSaveFolder, clicking, clickLocations
    
    if (clicking)
        ToggleClicking()
    
    ; Changed file extension from .click to .rcl
    filePath := FileSelect("1", defaultSaveFolder, "Select a click list to load", "Regy Click List (*.rcl)")
    
    if (filePath = "")
        return
    
    try
    {
        ClickTable.Delete()
        clickLocations := []
        
        FileObj := FileOpen(filePath, "r")
        
        inLocationSection := false
        currentLocation := {}
        
        while !FileObj.AtEOF
        {
            line := FileObj.ReadLine()
            
            line := RTrim(line, "`r`n")
            
            if (SubStr(line, 1, 1) = "[" && SubStr(line, -1) = "]")
            {
                if (SubStr(line, 1, 9) = "[Location")
                {
                    if (inLocationSection && currentLocation.HasProp("x") && currentLocation.HasProp("y") 
                        && currentLocation.HasProp("delay") && currentLocation.HasProp("clickType"))
                    {
                        if (!currentLocation.HasProp("winTitle"))
                            currentLocation.winTitle := "Unknown Window"
                        if (!currentLocation.HasProp("winHwnd"))
                            currentLocation.winHwnd := ""
                            
                        clickLocations.Push(currentLocation)
                    }
                    
                    inLocationSection := true
                    currentLocation := {}
                }
                else
                {
                    inLocationSection := false
                }
                
                continue
            }
            
            if (inLocationSection)
            {
                if (InStr(line, "="))
                {
                    parts := StrSplit(line, "=", , 2)
                    key := parts[1]
                    value := parts[2]
                    
                    if (key = "X")
                        currentLocation.x := value
                    else if (key = "Y")
                        currentLocation.y := value
                    else if (key = "Delay")
                        currentLocation.delay := value
                    else if (key = "ClickType")
                        currentLocation.clickType := value
                    else if (key = "WindowTitle")
                        currentLocation.winTitle := value
                    else if (key = "WindowID")
                        currentLocation.winHwnd := value
                    else if (key = "LoopType")
                        currentLocation.loopType := value
                    else if (key = "LoopCount")
                        currentLocation.loopCount := value
                }
            }
        }
        
        if (inLocationSection && currentLocation.HasProp("x") && currentLocation.HasProp("y") 
            && currentLocation.HasProp("delay") && currentLocation.HasProp("clickType"))
        {
            if (!currentLocation.HasProp("winTitle"))
                currentLocation.winTitle := "Unknown Window"
            if (!currentLocation.HasProp("winHwnd"))
                currentLocation.winHwnd := ""
                
            clickLocations.Push(currentLocation)
        }
        
        FileObj.Close()
        
        for i, location in clickLocations
        {
            ; Format for display with new column structure
            if (location.clickType = "LOOP START") {
                ClickTable.Add("", "", "Loop", "Start", location.loopCount, "")
            } 
            else if (location.clickType = "LOOP END") {
                ClickTable.Add("", "", "Loop", "End", "", "")
            }
            else {
                posDisplay := location.x "," location.y
                ClickTable.Add("", "", location.clickType, posDisplay, location.delay, location.winHwnd)
            }
        }
        
        currentIndex := 1
        
        ; Don't save loaded locations back to INI since we're just loading them
        ; SaveClickLocations()
        
        UpdateListNumbers()
        
        MsgBox("Click list loaded with " clickLocations.Length " click locations")
        StatusBar.SetText("Loaded " clickLocations.Length " locations from: " filePath)
    }
    catch as e
    {
        MsgBox("Error loading configuration: " e.Message)
        StatusBar.SetText("Error loading configuration")
    }
}

SaveConfigForExit(*)
{
    global defaultSaveFolder, saveCompleted
    
    SaveConfigPathDisplay.Value := defaultSaveFolder
    SaveConfigNameInput.Value := "MyClickList"
    
    ; Override the save button action
    SaveConfigSaveBtn.OnEvent("Click", SaveConfigurationFileAndExit)
    
    ; Also handle if user cancels
    SaveConfigCancelBtn.OnEvent("Click", CancelSaveAndExit)
    SaveConfigDialog.OnEvent("Close", CancelSaveAndExit)
    
    SaveConfigDialog.Show()
}

SaveConfigurationFileAndExit(*)
{
    global saveCompleted
    
    savePath := SaveConfigPathDisplay.Value
    configName := SaveConfigNameInput.Value
    
    if (configName = "")
    {
        MsgBox("Please enter a configuration name")
        return
    }
    
    configName := RegExReplace(configName, "[\\/:*?`"<>|]", "_")
    
    ; Changed file extension from .click to .rcl
    filePath := savePath "\" configName ".rcl"
    
    try
    {
        FileObj := FileOpen(filePath, "w")
        
        FileObj.WriteLine("[AutoClickerList]")
        FileObj.WriteLine("Version=1.1")
        FileObj.WriteLine("SaveDate=" A_Now)
        FileObj.WriteLine("LocationCount=" clickLocations.Length)
        
        for i, location in clickLocations
        {
            FileObj.WriteLine("[Location" i "]")
            FileObj.WriteLine("X=" location.x)
            FileObj.WriteLine("Y=" location.y)
            FileObj.WriteLine("Delay=" location.delay)
            FileObj.WriteLine("ClickType=" location.clickType)
            
            if (location.HasProp("loopType") && location.loopType)
                FileObj.WriteLine("LoopType=" location.loopType)
                
            if (location.HasProp("loopCount") && location.loopCount)
                FileObj.WriteLine("LoopCount=" location.loopCount)
                
            FileObj.WriteLine("WindowTitle=" location.winTitle)
            FileObj.WriteLine("WindowID=" location.winHwnd)
        }
        
        FileObj.Close()
        
        MsgBox("Click list saved to " filePath)
        SaveConfigDialog.Hide()
        StatusBar.SetText("Click list saved to: " filePath)
        
        ; Mark save as completed
        saveCompleted := true
        
        ; Now exit the application
        SaveSettings()
        ExitApp()
    }
    catch as e
    {
        MsgBox("Error saving configuration: " e.Message)
        StatusBar.SetText("Error saving configuration")
        
        ; Ask if the user wants to exit anyway
        result := MsgBox("Error saving. Exit anyway?", "Save Error", "YesNo")
        if (result = "Yes") {
            SaveSettings()
            ExitApp()
        }
    }
}

CancelSaveAndExit(*)
{
    SaveConfigDialog.Hide()
    
    ; Ask if the user wants to exit without saving
    result := MsgBox("Exit without saving?", "Confirm Exit", "YesNo")
    if (result = "Yes") {
        SaveSettings()
        ExitApp()
    }
}

; ==============================================================================
; MAIN EXECUTION
; ==============================================================================
MainWindow.OnEvent("Close", CloseMainWindow)

CloseMainWindow(*)
{
    global clickLocations, isClosing, saveCompleted, clicking
    
    if (clicking) {
        result := MsgBox("Clicking is active. Exit anyway?", "Confirm Exit", "YesNo")
        if (result != "Yes")
            return
    }
    
    ; Set closing flag
    isClosing := true
    saveCompleted := false
    
    ; Only ask to save if there are locations to save
    if (clickLocations.Length > 0) {
        result := MsgBox("Would you like to save your click list before exiting?", "Save Click List", "YesNo")
        if (result = "Yes") {
            SaveConfigForExit()
            return ; Don't exit yet - wait for save to complete
        }
    }
    
    ; If no save needed or user declined to save, exit now
    SaveSettings() ; Still save settings like hotkeys
    ExitApp()
}

; Initialize
if !DirExist(defaultSaveFolder)
    DirCreate(defaultSaveFolder)

LoadSavedSettings() ; Only load settings, not click list
SetupHotkeys()

MainWindow.Show()