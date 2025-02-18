Dim userChoice, i, objShell, scriptPath

Set objShell = CreateObject("WScript.Shell")
scriptPath = WScript.ScriptFullName ' Get the current script file path

' Show the main dialog box with Yes (Continue) and No (Quit) buttons
userChoice = MsgBox("Do you want to Continue or Quit?", vbYesNo + vbQuestion, "Choose an Option")

If userChoice = vbYes Then
    ' If the user clicks "Yes" (Continue)
    MsgBox "Have a good time!", vbInformation, "Goodbye!"
    WScript.Sleep 5000 ' Wait 5 seconds and exit
Else
    ' If the user clicks "No" (Quit), open 4 new instances of this script
    For i = 1 To 4
        objShell.Run "wscript.exe """ & scriptPath & """", 0, False
    Next

    ' Wait for 5 seconds before killing all tasks
    WScript.Sleep 30000  ' Delay for 30 seconds

    ' Kill all instances of wscript.exe to close all windows at once
    objShell.Run "taskkill /F /IM wscript.exe", 0, False
End If
