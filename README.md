# TTS_selected_text_globally_in_Windows
TTS (Text-to-Speech) any selected text globally in Windows.
A lightweight, "no-software-required" solution to make Windows read any selected text aloud using a keyboard shortcut. Works in Chrome, Word, PDFs, and more.

## The Story (Why this exists)
Windows has built-in TTS, but lacks a global "Speak Selection" hotkey. This project uses a native VBScript wrapper to bridge that gap. We utilize the 32-bit Windows Script Host to ensure maximum compatibility with system voices.

---

## 1. Installation

### Step A: Save the Scripts
1. Create a folder at `C:\Scripts\`.
2. Download `SpeakText.vbs` from this repo and place it there.

### Step B: Create the Global Hotkey
Windows only monitors the **Desktop** and **Start Menu** for shortcut keys.
1. hold Right-click and drag `SpeakText.vbs` > **Create Shortcut Here**.
2. Move that Shortcut to your **Desktop**.
3. Right-click the Shortcut > **Properties**.
4. **The Secret Sauce:** Change the **Target** field to:
   `C:\Windows\SysWOW64\wscript.exe "C:\Scripts\SpeakText.vbs"`
5. Click the **Shortcut key** box and press `Ctrl + Alt + S`. I used the ***** key in the Numpad.
6. Set **Run** to `Minimized`.

---

## 2. How it Works
When you press the hotkey:
1. It sends `Alt+Tab` to refocus your previous app.
2. It sends `Ctrl+C` to copy your selection.
3. It uses the `SAPI.SpVoice` object to read the clipboard buffer.

---

## 3. Troubleshooting (Read if it fails!)

### "No text detected" or "Clipboard empty"
* **The Fix:** Increase the `WScript.Sleep` values in the script. Modern apps like Chrome sometimes need 300-500ms to process a "Copy" command.
* **Focus:** Ensure you are not running an "Admin" app (like Task Manager) as the active window; Windows blocks script-sent keys to Admin windows for security.

### "Object Required" or Voice not playing
* **The Fix:** Ensure your shortcut is pointing to `C:\Windows\SysWOW64\wscript.exe`. This forces the script to run in 32-bit mode, which is required for certain SAPI voice registrations.

### Script won't trigger
* **The Fix:** Move the shortcut to your **Desktop** or `shell:start menu`. Windows ignores hotkeys assigned to shortcuts buried in standard folders.

---

## 4. The Code
```vbs
Set objShell = CreateObject("WScript.Shell")
Set objVoice = CreateObject("SAPI.SpVoice")

' 1. Small delay to let you release the physical keys
WScript.Sleep 200

' 2. Force a Ctrl+C by "holding" the keys (using lowercase 'c' is often more reliable)
' This sends an Alt+Tab back and forth quickly to re-prime the focus
objShell.SendKeys "%{TAB}"
WScript.Sleep 300
objShell.SendKeys "^c"
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
```
