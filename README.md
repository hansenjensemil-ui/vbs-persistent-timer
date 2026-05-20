Requires NirCmd to run and launch "bubble-messages" 
- provides a clean way to give messages to user

 ================================================
 Persistent Timer Script for VBScript (timer.vbs) - CLEANED & REFACTORED
 ================================================
 Usage examples (run with cscript or wscript):
   cscript timer.vbs @12:45 "The timer has gone off"
   cscript timer.vbs 45m "Remember dentist"
   cscript timer.vbs 05:45 start notepad.exe
   cscript timer.vbs 2h start "C:\Program Files\Notepad++\notepad++.exe"

 Features:
 - Supports @hh:mm (absolute today/tomorrow if past)
 - Supports hh:mm (relative hours:minutes)
 - Supports ##s / ##m / ##h (relative seconds/minutes/hours)
 - Action: quoted message → system-modal MsgBox
 - Action: "start" + executable/path → launch program
 - Full persistence in timers.txt (same folder as script)
 - Unique hash per timer
 - On every launch: checks unfinished (active) timers
   • Overdue timers: asks "Execute now?" (MsgBox)
   • Future timers: left untouched (logged)
   • Done timers kept as log
 - Uses WScript.Sleep for the new timer
 - Uses nircmd TrayTip bubbles for all status messages
 - MsgBox / Echo only used when really needed

Made by Jens Emil using Grok