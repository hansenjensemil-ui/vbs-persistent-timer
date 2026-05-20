' ================================================
' Persistent Timer Script for VBScript (timer.vbs) - CLEANED & REFACTORED
' ================================================
' Usage examples (run with cscript or wscript):
'   cscript timer.vbs @12:45 "The timer has gone off"
'   cscript timer.vbs 45m "Remember dentist"
'   cscript timer.vbs 05:45 start notepad.exe
'   cscript timer.vbs 2h start "C:\Program Files\Notepad++\notepad++.exe"
'
' Features:
' - Supports @hh:mm (absolute today/tomorrow if past)
' - Supports hh:mm (relative hours:minutes)
' - Supports ##s / ##m / ##h (relative seconds/minutes/hours)
' - Action: quoted message → system-modal MsgBox
' - Action: "start" + executable/path → launch program
' - Full persistence in timers.txt (same folder as script)
' - Unique hash per timer
' - On every launch: checks unfinished (active) timers
'   • Overdue timers: asks "Execute now?" (MsgBox)
'   • Future timers: left untouched (logged)
'   • Done timers kept as log
' - Uses WScript.Sleep for the new timer
' - Uses nircmd TrayTip bubbles for all status messages
' - MsgBox / Echo only used when really needed

Option Explicit

' ================================================
' Helper: IIf (VBScript does not have built-in IIf)
' ================================================
Function IIf(expr, truepart, falsepart)
    If expr Then
        IIf = truepart
    Else
        IIf = falsepart
    End If
End Function

On Error Resume Next
On Error GoTo 0

' ================================================
Dim fso, shell, timerFile, args
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Timer file in the same folder as the script
timerFile = fso.GetParentFolderName(WScript.ScriptFullName) & "\timers.txt"

Dim qm: qm = chr(34)

Set args = WScript.Arguments

' ================== MAIN ==================
' 1. Always check/recover unfinished timers on startup
CheckPendingTimers

' 2. If arguments supplied → set a new timer
If args.Count > 0 Then
    ProcessNewTimer
Else
    ' Silent on normal startup with no new timer (only bubbles if overdue)
End If

WScript.Quit

' ================================================
' FUNCTION: Show non-intrusive bubble via nircmd
' ================================================
Function ShowBubble(title, text)
    On Error Resume Next
    NirMsg title, text, "shell32.dll,-16741", 20000
End Function

' ================================================
' FUNCTION: Check and handle unfinished timers
' ================================================
Function CheckPendingTimers()
    If Not fso.FileExists(timerFile) Then
        CheckPendingTimers = True
        Exit Function
    End If
    
    Dim lines, line, fields
    Dim hash, targetStr, typ, data, status
    Dim targetDate, response
    Dim newContent : newContent = ""
    Dim hasChanges : hasChanges = False
    
    lines = Split(fso.OpenTextFile(timerFile, 1).ReadAll, vbCrLf)
    
    Dim i
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" And Left(line, 1) <> "#" Then
            fields = Split(line, "|")
            If UBound(fields) >= 4 Then
                hash = fields(0)
                targetStr = fields(1)
                typ = fields(2)
                data = Replace(fields(3), "¦", "|")   ' restore original | if any
                status = fields(4)
                
                If status = "active" Then
                    On Error Resume Next
                    targetDate = CDate(targetStr)
                    On Error GoTo 0
                    
                    If Err.Number = 0 And targetDate <= Now Then
                        ' Overdue / due timer → ask user (MsgBox as requested)
                        response = MsgBox("Unfinished timer (ID: " & hash & ")" & vbCrLf & vbCrLf & _
                                          "Scheduled: " & targetDate & vbCrLf & _
                                          IIf(typ = "msg", "Message: ", "Launch: ") & data & vbCrLf & vbCrLf & _
                                          "Execute this action NOW?", _
                                          vbYesNo + vbQuestion + vbSystemModal, "Timer Overdue")
                        
                        If response = vbYes Then
                            ExecuteAction typ, data
                        End If
                        
                        ' Mark as done (whether executed or not)
                        newContent = newContent & hash & "|" & targetStr & "|" & typ & "|" & fields(3) & "|done" & vbCrLf
                        hasChanges = True
                    Else
                        ' Future timer → keep unchanged
                        newContent = newContent & line & vbCrLf
                    End If
                Else
                    ' Already done / cancelled → keep as log
                    newContent = newContent & line & vbCrLf
                End If
            Else
                newContent = newContent & line & vbCrLf
            End If
        Else
            newContent = newContent & line & vbCrLf
        End If
    Next
    
    If hasChanges Then
        Dim ts : Set ts = fso.CreateTextFile(timerFile, True)
        ts.Write newContent
        ts.Close
    End If
    CheckPendingTimers = True
End Function

' ================================================
' FUNCTION: Parse args and set a new timer
' ================================================
Function ProcessNewTimer()
    Dim timeSpec : timeSpec = Trim(args(0))
    Dim typ, data, targetDate, hash, targetStr
    
    ' Determine action type and data
    If args.Count > 1 Then
		If LCase(Trim(args(1))) = "start" Then
			typ = "start"
			data = ""
			Dim i
			For i = 2 To args.Count - 1
				data = data & args(i) & " "
			Next
			data = Trim(data)
		Else
			typ = "msg"
			data = ""
			For i = 1 To args.Count - 1
				data = data & args(i) & " "
			Next
			data = Trim(data)
		End If      
    End If
    
    'If data = "" Then
    '    MsgBox "Error: No action (message or program) specified.", vbCritical, "Timer Error"
    '    WScript.Quit 1
    'End If
    
    ' Calculate target datetime
    targetDate = GetTargetTime(timeSpec)
    If IsNull(targetDate) Then
        MsgBox "Error: Invalid time format: " & timeSpec, vbCritical, "Timer Error"
        WScript.Quit 1
    End If
    
    ' Unique hash (GUID)
    hash = GenerateUniqueHash()
    
    ' Store full ISO-like datetime
    targetStr = Year(targetDate) & "-" & Right("0" & Month(targetDate), 2) & "-" & Right("0" & Day(targetDate), 2) & " " & _
                Right("0" & Hour(targetDate), 2) & ":" & Right("0" & Minute(targetDate), 2) & ":" & Right("0" & Second(targetDate), 2)
    
    ' Append to persistent file
    AppendTimer hash, targetStr, typ, data
    
    ' Show bubble instead of console echo
    ShowBubble "Timer Scheduled", "Will trigger at " & targetDate & vbCrLf & _
               IIf(typ = "msg", "Message: ", "Launch: ") & data
    
    ' Sleep until due
    Dim sleepMs : sleepMs = DateDiff("s", Now, targetDate) 
    If sleepMs > 0 Then
        If sleepMs > 2147483647 Then sleepMs = 2147483647
        NirMsg "TimerScript", "Sleeping until " & targetDate & " (" & sleepMs & " s)", "shell32.dll,-16741", 20000
        WScript.Sleep sleepMs*1000
    End If
    
    ' Show completion bubble
    NirMsg "TimerScript done!", "Timer done", "shell32.dll,-16741", 35000

    ' Execute action
    ExecuteAction typ, data
    
    ' Mark as completed
    MarkTimerDone hash
    
    ProcessNewTimer = True
End Function

' ================================================
' FUNCTION: Parse time specifier → target Date
' ================================================
Function GetTargetTime(timeSpec)
    Dim t, parts, h, m, val
    timeSpec = Trim(timeSpec)
    
    On Error Resume Next
    
    If Left(timeSpec, 1) = "@" Then
        ' Absolute @hh:mm (today, or tomorrow if already passed)
        t = Mid(timeSpec, 2)
        GetTargetTime = Date + TimeValue(t)
        If GetTargetTime < Now Then GetTargetTime = GetTargetTime + 1
    ElseIf InStr(timeSpec, ":") > 0 Then
        ' Relative hh:mm
        parts = Split(timeSpec, ":")
        h = CInt(parts(0))
        m = CInt(parts(1))
        GetTargetTime = DateAdd("n", h * 60 + m, Now)
    ElseIf LCase(Right(timeSpec, 1)) = "s" Then
        val = CLng(Left(timeSpec, Len(timeSpec) - 1))
        GetTargetTime = DateAdd("s", val, Now)
    ElseIf LCase(Right(timeSpec, 1)) = "m" Then
        val = CLng(Left(timeSpec, Len(timeSpec) - 1))
        GetTargetTime = DateAdd("n", val, Now)
    ElseIf LCase(Right(timeSpec, 1)) = "h" Then
        val = CLng(Left(timeSpec, Len(timeSpec) - 1))
        GetTargetTime = DateAdd("h", val, Now)
    Else
        GetTargetTime = Null
    End If
    
    If Err.Number <> 0 Then GetTargetTime = Null
    On Error GoTo 0
End Function

' ================================================
' FUNCTION: Unique hash (GUID)
' ================================================
Function GenerateUniqueHash()
    Dim tlib
    Set tlib = CreateObject("Scriptlet.TypeLib")
    GenerateUniqueHash = Left(tlib.GUID, 38)
    Set tlib = Nothing
End Function

' ================================================
' FUNCTION: Append new timer to file
' ================================================
Function AppendTimer(hash, targetStr, typ, data)
    Dim ts
    Set ts = fso.OpenTextFile(timerFile, 8, True)
    ts.WriteLine hash & "|" & targetStr & "|" & typ & "|" & Replace(data, "|", "¦") & "|active"
    ts.Close
    AppendTimer = True
End Function

' ================================================
' FUNCTION: Mark a timer as done
' ================================================
Function MarkTimerDone(hash)
    If Not fso.FileExists(timerFile) Then
        MarkTimerDone = True
        Exit Function
    End If
    
    Dim content, lines, line, fields, newContent
    newContent = ""
    
    content = fso.OpenTextFile(timerFile, 1).ReadAll
    lines = Split(content, vbCrLf)
    
    Dim i
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            fields = Split(line, "|")
            If UBound(fields) >= 0 And fields(0) = hash Then
                fields(4) = "done"
                line = Join(fields, "|")
            End If
            newContent = newContent & line & vbCrLf
        End If
    Next
    
    fso.CreateTextFile(timerFile, True).Write newContent
    MarkTimerDone = True
End Function

' ================================================
' FUNCTION: Execute the scheduled action
' ================================================
Function ExecuteAction(typ, data)
    If typ = "msg" Then
        ' System modal MsgBox (as originally requested)
        Msgbox "Timer done!", vbOKOnly + vbSystemModal + vbInformation, "Timer done!"
    ElseIf typ = "start" Then
        ' Launch executable
        shell.Run chr(34) & data & chr(34), 1, False
    End If
    ExecuteAction = True
End Function

' ================================================
' FUNCTION: NirMsg - nircmd trayballoon wrapper
' ================================================
Function NirMsg(title, tekst, icon, timeout)
    qm = chr(34)
    shell.run "nircmd trayballoon " & qm & title & qm & " " & qm & tekst & qm & " " & qm & icon & qm & " " & timeout
    NirMsg = True
End Function

'Made by Jens Emil using Grok