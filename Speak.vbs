Set objShell = CreateObject("WScript.Shell")
Set objVoice = CreateObject("SAPI.SpVoice")

' 1. Small delay to let you release the physical keys
WScript.Sleep 200

' 2. Force a Ctrl+C by "holding" the keys (using lowercase 'c' is often more reliable)
' This sends an Alt+Tab back and forth quickly to re-prime the focus
objShell.SendKeys "%{TAB}"
WScript.Sleep 300
objShell.SendKeys "^c"

' 3. Give the clipboard a significant window to update
WScript.Sleep 400

' 4. Access the clipboard
Set objHTML = CreateObject("htmlfile")
selectedText = objHTML.ParentWindow.ClipboardData.GetData("text")

If Not IsNull(selectedText) And selectedText <> "" Then
    ' Optional: objVoice.Rate = 1 ' Uncomment to speed up the voice
    objVoice.Speak selectedText
End If