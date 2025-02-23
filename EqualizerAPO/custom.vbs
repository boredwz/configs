Dim classHelp: Set classHelp = New HelpMessage
Dim classPresetSystem: Set classPresetSystem = New Preset
Dim classPresetBaseus: Set classPresetBaseus = New Preset
Dim varConfigPath, varConfigDir, classPresetsArray
Dim varSetPresetSystem, varSetPresetBaseus, varShortcutsAdd, varShortcutsDel ' commands
Call Initial()
If varSetPresetSystem Then
    Call Set_PresetSystem()
ElseIf varSetPresetBaseus Then
    Call Set_PresetBaseus()
ElseIf varShortcutsAdd Then
    Call Set_Shortcuts(True, classPresetsArray)
ElseIf varShortcutsDel Then
    Call Set_Shortcuts(False, classPresetsArray)
Else
    Call MsgBox(classHelp.GetTextWarningArgsWrong(WScript.Arguments.Item(0)), vbExclamation+vbOKOnly, classHelp.Title)
End If
WScript.Quit





' Initial checks (Only one parameter, /help) and setup (Set Variables).
Sub Initial()
    Call Set_Variables()
    Dim choice
    If WScript.Arguments.Count = 0 Then
        If IsConfigInstalled(varConfigDir) Then
            Call classHelp.Show()
        Else
            choice = MsgBox(classHelp.GetTextInstallPrompt(), vbQuestion+vbYesNo, classHelp.Title)
            If choice = vbYes Then Call Install_Config()
        End If
        WScript.Quit
    ElseIf WScript.Arguments.Count > 1 Then
        Call MsgBox(classHelp.GetTextWarningArgsMultiple, vbExclamation+vbOKOnly, classHelp.Title)
        WScript.Quit
    End If
    Dim arr_help_pattern
    arr_help_pattern = Array("/help","/h","/?","--help","--h","--?","-help","-h","-?","?")
    If Check_Argument(arr_help_pattern) Then
        Call classHelp.Show()
        WScript.Quit
    End If
    If Not IsConfigInstalled(varConfigDir) Then
        choice = MsgBox(classHelp.GetTextInstallPrompt(), vbQuestion+vbYesNo, classHelp.Title)
        If choice = vbYes Then Call Install_Config()
        WScript.Quit
    End If
End Sub

Sub Set_Variables()
    ' command-line arguments: read and check
    Dim presetparam: presetparam = Array("preset","set","mode")
    Dim system: system = Array("system")
    Dim baseus: baseus = Array("baseus","baseus_white")
    Dim shortcutparam: shortcutparam = Array("startmenushortcuts","startmenushortcut","startmenu","start","shortcuts","shortcut","links","link","lnks","lnk")
    Dim shortcutAdd: shortcutAdd = Array("add","create","set","+")
    Dim shortcutDel: shortcutDel = Array("delete","remove","del","-")
    varSetPresetSystem = Check_Argument_Named(presetparam, system)
    varSetPresetBaseus = Check_Argument_Named(presetparam, baseus)
    varShortcutsAdd = Check_Argument_Named(shortcutparam, shortcutAdd)
    varShortcutsDel = Check_Argument_Named(shortcutparam, shortcutDel)

    ' presets
    With classPresetSystem
        .Name = "System speakers"
        .Path = "custom\SYSTEM_LENOVO_80WK.txt"
        .ShortcutIconPath = "%WINDIR%\system32\DDORes.dll,14"
        .ShortcutTargetArguments = "/preset:system"
    End With
    With classPresetBaseus
        .Name = "Baseus Wired White"
        .Path = "custom\3.5_BASEUS_WHITE.txt"
        .ShortcutIconPath = "%WINDIR%\system32\DDORes.dll,6"
        .ShortcutTargetArguments = "/preset:baseus"
    End With
    classPresetsArray = Array(classPresetSystem, classPresetBaseus)
    
    ' config path
    Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
    varConfigDir = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%\EqualizerAPO\config")
    varConfigPath = varConfigDir & "\custom.txt"
End Sub

Sub Install_Config()
    Dim objShell: Set objShell = CreateObject("WScript.Shell")
    Dim objShellApp: Set objShellApp = CreateObject("Shell.Application")
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim parent: parent = objFSO.GetParentFolderName(WScript.ScriptFullName)
    Dim customDir: customDir = objFSO.BuildPath(parent, "custom")
    Dim files: files = "custom.vbs custom.txt config.txt"
    Dim valid: valid = objFSO.FolderExists(customDir)
    For Each i In Split(files, " ")
        Dim j: j = objFSO.BuildPath(parent, i)
        valid = valid And objFSO.FileExists(j)
    Next
    If Not valid Then
        Call MsgBox(classHelp.GetTextWarningInstallWrong, vbExclamation+vbOKOnly, classHelp.Title)
        WScript.Quit
    End If
    Dim cmd: cmd = "ROBOCOPY """ & customDir & """ """ & varConfigDir & "\custom"" /E"
    Dim cmdd: cmdd = "ROBOCOPY """ & parent & """ """ & varConfigDir & """ " & files    
    Call objShell.Run("cmd /c " & cmd, 0, True)
    Call objShell.Run("cmd /c " & cmdd, 0, True)
    Call objShellApp.ShellExecute(varConfigDir & "\custom.vbs", "/shortcuts:add", "", "", 1)
End Sub

' 0-100, even numbers only.
Sub Set_Volume(value)
    Dim objShell: Set objShell = CreateObject("WScript.Shell")
	
	For i = 1 to 100
		objShell.SendKeys(chr(&hAE)) ' volume down
	Next
	
	For i = 1 to Int(value / 2)
		objShell.SendKeys(chr(&hAF)) ' volume up
	Next
End Sub





Sub Set_PresetSystem()
    Dim result: result = _
        Enable_Preset(varConfigPath, classPresetSystem) And _
        Disable_Preset(varConfigPath, classPresetBaseus)
    Call Set_Volume(14)
    If Not result Then Call MsgBox("Set_PresetSystem: " & result, vbExclamation+vbOKOnly, classHelp.Title)
End Sub

Sub Set_PresetBaseus()
    Dim result: result = _
        Enable_Preset(varConfigPath, classPresetBaseus) And _
        Disable_Preset(varConfigPath, classPresetSystem)
    Call Set_Volume(32)
    If Not result Then Call MsgBox("Set_PresetBaseus: " & result, vbExclamation+vbOKOnly, classHelp.Title)
End Sub





' Return True if the shortcuts created (or removed) successfully.
Function Set_Shortcuts(create, presets)
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    Dim result: result = True

    For Each i in presets
        If create Then
            result = result And Shortcut_Create( _
                i.ShortcutTarget, _
                i.ShortcutTargetArguments, _
                i.ShortcutWorkDir, _
                i.ShortcutPath, _
                i.ShortcutIconPath _
            )
        Else
            On Error Resume Next
            Call objFSO.DeleteFile(i.ShortcutPath, True)
            result = result And (Err.Number = 0)
            Err.Clear
            On Error Goto 0
        End If
    Next

    Shortcuts = result
End Function

Function Enable_Preset(config, preset)
    Dim line: line = "Include: " & preset.Path
    Dim strMatch: strMatch = "(\# )?" & "# " & RegexEscape(line)
    Dim strReplace: strReplace = line
    Enable_Preset = RegexReplaceInFile(config, strMatch, strReplace)
End Function

Function Disable_Preset(config, preset)
    Dim line: line = "Include: " & preset.Path
    Dim strMatch: strMatch = "(\# )?" & RegexEscape(line)
    Dim strReplace: strReplace = "# " & line
    Disable_Preset = RegexReplaceInFile(config, strMatch, strReplace)
End Function

Function IsConfigInstalled(dir)
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim result: result = objFSO.FolderExists(objFSO.BuildPath(dir, "custom"))
    For Each i In Array("custom.vbs","custom.txt","config.txt")
        Dim j: j = objFSO.BuildPath(dir, i)
        result = result And objFSO.FileExists(j)
    Next
    IsConfigInstalled = result
End Function





' Return True if the argument value matches the pattern.
Function Check_Argument(name_pattern)
    ' string -> array
    Dim arr_names
    arr_names = name_pattern
    If Not IsArray(name_pattern) Then arr_names = Array(name_pattern)

    Check_Argument = False
    For Each argument in WScript.Arguments
        Dim match
        match = False

        If Not IsEmpty(argument) Then
            For Each pattern in arr_names
                match = (StrComp(argument, pattern, 1) = 0)
                If match Then Exit For
            Next
        End If

        If match Then
            Check_Argument = True
            Exit For
        End If
    Next
End Function

' (Using named args) Return True if the argument value matches the pattern.
Function Check_Argument_Named(name_pattern, value_pattern)
    ' string -> array
    Dim arr_names, arr_values
    arr_names = name_pattern
    arr_values = value_pattern
    If Not IsArray(name_pattern) Then arr_names = Array(name_pattern)
    If Not IsArray(value_pattern) Then arr_values = Array(value_pattern)

    Check_Argument_Named = False
    For Each parameter in arr_names
        Dim argument_value, match

        argument_value = WScript.Arguments.Named.Item(parameter)
        match = False

        If Not IsEmpty(argument_value) Then
            For Each pattern in arr_values
                match = (StrComp(argument_value, pattern, 1) = 0)
                If match Then Exit For
            Next
        End If

        If match Then
            Check_Argument_Named = True
            Exit For
        End If
    Next
End Function

' Return True if .lnk is created successfully.
Function Shortcut_Create(target, arguments, workdir, lnk, lnk_icon)
    Dim shortcut_path, target_expanded
    Dim objShell: Set objShell = CreateObject("WScript.Shell")
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Shortcut_Create = False
    If IsEmpty(target) Or IsNull(target) Then Exit Function

    ' Set the shortcut path based on the target path
    shortcut_path = lnk
    If IsEmpty(lnk) Or IsNull(lnk) Then
        shortcut_path = objFSO.BuildPath(objFSO.GetParentFolderName(target), objFSO.GetBaseName(target) & ".lnk")
    End If
    shortcut_path = objShell.ExpandEnvironmentStrings(shortcut_path)

    If objFSO.FileExists(shortcut_path) Then Exit Function

    Dim objShortcut: Set objShortcut = objShell.CreateShortcut(shortcut_path)
    objShortcut.TargetPath = target
    If (Not IsEmpty(arguments)) And (Not IsNull(arguments)) Then objShortcut.Arguments = arguments
    If (Not IsEmpty(workdir)) And (Not IsNull(workdir)) Then objShortcut.WorkingDirectory = workdir
    If (Not IsEmpty(lnk_icon)) And (Not IsNull(lnk_icon)) Then objShortcut.IconLocation = lnk_icon
    On Error Resume Next
    objShortcut.Save
    If Err.Number <> 0 Then
        WScript.Echo "Shortcut_Create error." & vbNewLine & " target: " & target & vbNewLine & " info: " & Err.Description
        Err.Clear
    End If
    On Error Goto 0
    If objFSO.FileExists(shortcut_path) Then Shortcut_Create = True
End Function

' Return True if string is replaced successfully.
Function RegexReplaceInFile(filePath, strMatch, strReplace)
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile: Set objFile = objFSO.OpenTextFile(filePath, 1)
    Dim strContent: strContent = objFile.ReadAll
    objFile.Close
    RegexReplaceInFile = True

    Dim objRegExp: Set objRegExp = New RegExp
    With objRegExp
        .Pattern = strMatch
        .Global = True
        .IgnoreCase = True
    End With
    
    strContent = objRegExp.Replace(strContent, strReplace)
    Set objFile = objFSO.OpenTextFile(filePath, 2)
    On Error Resume Next
    objFile.Write strContent
    If Err.Number <> 0 Then RegexReplaceInFile = False
    Err.Clear
    objFile.Close
    If Err.Number <> 0 Then RegexReplaceInFile = False
    On Error Goto 0
End Function

' Return text with regex chars escaped.
Function RegexEscape(text)
    Dim txt, symbol
    txt = text
    symbol = Array("\",".","*","&","^","%","$","#","@","!",":","<",">","?","/","(",")","[","]","{","}","+","|")
    For i=LBound(symbol) To UBound(symbol)
        txt = Replace(txt, symbol(i), "\" & symbol(i))
    Next
    RegexEscape = txt
End Function





Class Preset
    Public Path, ShortcutPath, ShortcutIconPath, ShortcutWorkDir, ShortcutTarget, ShortcutTargetArguments
    Private prName, prStartMenuPath
    Public Property Get Name
        Name = prName
    End Property
    Public Property Let Name(value)
        prName = value
        ShortcutPath = prStartMenuPath & prName & " (EQ APO).lnk"
    End Property

    Public Function GetExpanded(path)
        Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
        GetExpanded = objShell.ExpandEnvironmentStrings(path)
    End Function

    ' Defaults
    Private Sub Class_Initialize()
        ShortcutTarget = "%PROGRAMFILES%\EqualizerAPO\config\custom.vbs"
        prStartMenuPath = GetExpanded("%APPDATA%\Microsoft\Windows\Start Menu\Programs\")
    End Sub
End Class

Class HelpMessage
    Private prText, prTitle
    Public Property Get Text
        Text = prText
    End Property
    Public Property Get Title
        Title = prTitle
    End Property
    
    Public Function GetTextWarningArgsWrong(arguments)
        GetTextWarningArgsWrong = _
            "Wrong arguments:" & vbNewLine & _
            WScript.Arguments.Item(0) & vbNewLine & vbNewLine & _
            "Use /help for more information."
    End Function

    Public Function GetTextInstallPrompt()
        GetTextInstallPrompt = _
            "Do you want to install custom config?" & vbNewLine & vbNewLine & _
            "<!> NOTE: This will replace your current Equalizer APO config!"
    End Function
    
    Public Function GetTextWarningArgsMultiple()
        GetTextWarningArgsMultiple = _
            "Only one parameter is supported." & vbNewLine & vbNewLine & _
            "Use /help for more information."
    End Function

    Public Function GetTextWarningInstallWrong()
        GetTextWarningInstallWrong = _
            "Configuration files not found, place this file in the same dir with:." & vbNewLine & _
            "FILES: config.txt, custom.txt, custom.vbs" & _
            "DIR:   custom"
    End Function

    Public Sub Show()
        MsgBox prText, vbInformation+vbOKOnly, prTitle
    End Sub
    
    Private Sub Reset()
        prTitle = "Equalizer APO custom config (boredwz)"
        prText = _
            "USAGE:" & vbNewLine & _
            "    ""cscript.exe //nologo custom.vbs [/Parameter[:Value]]""" & vbNewLine & _
            "PARAMETERS:" & vbNewLine & _
            "    /preset:[PRESET_NAME]  - Set custom preset" & vbNewLine & _
            "    /shortcuts:[ADD/DEL]   - Create or remove Start Menu shortcuts" & vbNewLine & _
            "PRESETS:" & vbNewLine & _
            "    ""System""  - Reset to the system" & vbNewLine & _
            "    ""Baseus""  - Baseus Wired White"
    End Sub

    Private Sub Class_Initialize()
        Reset()
    End Sub
End Class