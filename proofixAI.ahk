/*
ProofixAI

Author: Arshit Vaghasiya | Geek Updates
Date: 13-August-2025
More tech tips and tricks: https://geek-updates.com/
For the tutorial and more details, visit: http://bit.ly/proofixai

WHAT IT DOES:
- Select text, press Alt+P: Fix grammar/spelling instantly
- Alt+L: View edit history

CREDITS:
cJson.ahk - The first and only AutoHotkey JSON library to use embedded compiled C for high performance
Author: G33kDude
Source: https://github.com/G33kDude/cJson.ahk

Gemini API integration code by u/Laser_Made
Author: https://www.reddit.com/user/Laser_Made/
Source: https://www.reddit.com/r/AutoHotkey/comments/1ci2x6q/comment/l2hrijw/
*/

#Requires AutoHotkey v2.0+
#SingleInstance Force
#include JSON.ahk

TraySetIcon "logo.ico"

; Toast notification class for non-blocking messages
class Toast {
  static Show(message, duration := 3000) {
    ; Get mouse position for toast placement
    MouseGetPos(&mouseX, &mouseY)

    ; Create toast GUI
    toastGui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox +LastFound", "ProofixAI")
    toastGui.MarginX := 15
    toastGui.MarginY := 10
    toastGui.BackColor := "0x2D2D30"

    ; Add text control
    textCtrl := toastGui.Add("Text", "c0xFFFFFF", message)
    textCtrl.SetFont("s9", "Segoe UI")

    ; Position and show the toast
    toastGui.Show("x" . (mouseX + 15) . " y" . (mouseY + 15) . " AutoSize NoActivate")

    ; Auto-hide after duration
    if (duration > 0) {
      SetTimer(() => toastGui.Destroy(), -duration)
    }

    return toastGui
  }
}

; Global variables
logFile := A_AppData . "\ProofixAI_log.txt"
filePath := ".\geminiAPI.txt"
geminiAPIkey := FileRead(filePath, "UTF-8")

; Initialize log file with UTF-8 encoding
InitializeLog() {
  if (!FileExist(logFile)) {
    FileAppend("=== Text Proofreading Log ===`n", logFile, "UTF-8")
    FileAppend("Started: " . FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") . "`n`n", logFile, "UTF-8")
  }
}

; Simple proofreading function
ProofreadText(inputText, apikey) {
  prompt :=
    "Proofread the following text for grammar, spelling, and punctuation errors. Provide only the corrected text. Do not include any introductory or concluding remarks. Text: `"" .
    inputText . "`""
  strUrl := "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?key=" apikey

  ; Show thinking status
  Toast.Show("🤔 Thinking...")

  ; Log input
  LogSession(inputText, "")

  api := ComObject("MSXML2.XMLHTTP")

  try {
    api.Open("POST", strUrl, false)
    api.SetRequestHeader("Content-Type", "application/json")
    api.Send(JSON.Dump({ contents: [{ parts: [{ text: prompt }] }] }))

    while (api.readyState != 4) {
      Sleep(100)
    }

    if (api.status != 200) {
      Toast.Show("❌ Error occurred", 4000)
      return ""
    }

    ; Process response and replace text
    response := api.responseText
    correctedText := ProcessAndReplace(response, inputText)

    ; Log output
    LogSession(inputText, correctedText)

    ; Show completion status via Toast
    if (correctedText != inputText) {
      Toast.Show("✅ Text corrected", 2500)
    } else {
      Toast.Show("ℹ️ No changes needed", 2500)
    }

    return correctedText

  } catch Error as e {
    Toast.Show("❌ Connection failed", 4000)
    return ""
  }
}

; Process response and replace text in real-time
ProcessAndReplace(response, originalText) {
  try {
    ; Clean response
    cleanResponse := Trim(response, "[]`n`r `t")
    jsonObjects := SplitJSON(cleanResponse)

    ; Step 1: Delete the selected text (cursor is now at the position where text was)
    Send("{Delete}")

    ; Step 2: Process each chunk and append it at cursor position
    for index, jsonStr in jsonObjects {
      try {
        data := JSON.load(jsonStr)
        if (data.Has("candidates") && data["candidates"].Length > 0) {
          candidate := data["candidates"][1]
          if (candidate.Has("content") && candidate["content"].Has("parts") && candidate["content"]["parts"].Length > 0
          ) {
            newChunk := candidate["content"]["parts"][1]["text"]

            ; Simply append the new chunk at current cursor position
            SendText(newChunk)

            ; Show writing progress
            Toast.Show("✍️ Writing...", 1500)
          }
        }
      } catch {
        continue
      }
    }

    ; Build complete text for logging
    fullText := ""
    for index, jsonStr in jsonObjects {
      try {
        data := JSON.load(jsonStr)
        if (data.Has("candidates") && data["candidates"].Length > 0) {
          candidate := data["candidates"][1]
          if (candidate.Has("content") && candidate["content"].Has("parts") && candidate["content"]["parts"].Length > 0
          ) {
            fullText .= candidate["content"]["parts"][1]["text"]
          }
        }
      } catch {
        continue
      }
    }

    return fullText

  } catch {
    Toast.Show("❌ Processing failed", 4000)
    return ""
  }
}

; Split JSON helper
SplitJSON(jsonString) {
  jsonObjects := []
  currentObject := ""
  braceCount := 0
  inString := false
  escapeNext := false

  loop parse, jsonString {
    char := A_LoopField

    if (escapeNext) {
      escapeNext := false
      currentObject .= char
      continue
    }

    if (char == "\") {
      escapeNext := true
      currentObject .= char
      continue
    }

    if (char == '"') {
      inString := !inString
    }

    if (!inString) {
      if (char == "{") {
        braceCount++
      } else if (char == "}") {
        braceCount--
      }
    }

    currentObject .= char

    if (!inString && braceCount == 0 && currentObject != "") {
      currentObject := Trim(currentObject, " `t`n`r,")
      if (currentObject != "") {
        jsonObjects.Push(currentObject)
        currentObject := ""
      }
    }
  }

  return jsonObjects
}

; Modified logging function - writes new entries at the top
LogSession(inputText, outputText) {
  timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")

  ; Clean text to handle special characters
  cleanInput := CleanTextForLogging(inputText)

  ; Read existing log content
  existingContent := ""
  if (FileExist(logFile)) {
    existingContent := FileRead(logFile, "UTF-8")
  }

  ; Create new log entry
  if (outputText == "") {
    newEntry := "--- " . timestamp . " ---`n"
    newEntry .= "INPUT: " . cleanInput . "`n"
  } else {
    cleanOutput := CleanTextForLogging(outputText)
    newEntry := "OUTPUT: " . cleanOutput . "`n"
    if (inputText == outputText) {
      newEntry .= "RESULT: No changes`n`n"
    } else {
      newEntry .= "RESULT: Corrected`n`n"
    }
  }

  ; If this is a new session (input only), add separator and place at top
  if (outputText == "") {
    ; Find the position after the header (after the first empty line)
    headerEnd := InStr(existingContent, "`n`n")
    if (headerEnd > 0) {
      ; Insert new entry after header
      headerPart := SubStr(existingContent, 1, headerEnd + 1)
      restOfContent := SubStr(existingContent, headerEnd + 2)
      newContent := headerPart . newEntry . restOfContent
    } else {
      ; If no header found, just prepend
      newContent := newEntry . existingContent
    }
  } else {
    ; This is the output part, find the most recent INPUT entry and add OUTPUT after it
    lines := StrSplit(existingContent, "`n")
    newLines := []
    outputAdded := false

    for index, line in lines {
      newLines.Push(line)
      ; If we find an INPUT line that doesn't have an OUTPUT yet, add our OUTPUT
      if (!outputAdded && RegExMatch(line, "^INPUT:") && (index == lines.Length || !RegExMatch(lines[index + 1],
        "^OUTPUT:"))) {
        newLines.Push("OUTPUT: " . CleanTextForLogging(outputText))
        if (inputText == outputText) {
          newLines.Push("RESULT: No changes")
        } else {
          newLines.Push("RESULT: Corrected")
        }
        newLines.Push("")  ; Add empty line after complete entry
        outputAdded := true
      }
    }

    newContent := ""
    for index, line in newLines {
      newContent .= line
      if (index < newLines.Length)
        newContent .= "`n"
    }
  }

  ; Write the complete new content back to file
  FileDelete(logFile)
  FileAppend(newContent, logFile, "UTF-8")
}

; Clean text for proper logging - simplified approach
CleanTextForLogging(text) {
  if (text == "")
    return ""

  cleanText := text

  try {
    ; Replace common problematic Unicode characters using Chr() codes
    cleanText := StrReplace(cleanText, Chr(8220), '"')    ; " Left double quote
    cleanText := StrReplace(cleanText, Chr(8221), '"')    ; " Right double quote
    cleanText := StrReplace(cleanText, Chr(8216), "'")    ; ' Left single quote
    cleanText := StrReplace(cleanText, Chr(8217), "'")    ; ' Right single quote
    cleanText := StrReplace(cleanText, Chr(8212), "--")   ; — Em dash
    cleanText := StrReplace(cleanText, Chr(8211), "-")    ; – En dash
    cleanText := StrReplace(cleanText, Chr(8230), "...")  ; … Ellipsis
    cleanText := StrReplace(cleanText, Chr(8226), "*")    ; • Bullet point
    cleanText := StrReplace(cleanText, Chr(8482), "(TM)") ; ™ Trademark
    cleanText := StrReplace(cleanText, Chr(174), "(R)")   ; ® Registered
    cleanText := StrReplace(cleanText, Chr(169), "(C)")   ; © Copyright
  } catch Error as e {
    ; If cleaning fails, return original text
    return text
  }

  return cleanText
}

; Main hotkey - Alt+P
!p:: {
  ; Get selected text
  oldClipboard := A_Clipboard
  A_Clipboard := ""
  Send("^c")

  if (!ClipWait(1)) {
    Toast.Show("❌ No text selected", 3000)
    return
  }

  selectedText := A_Clipboard
  A_Clipboard := oldClipboard

  if (selectedText == "" || StrLen(Trim(selectedText)) < 1) {
    Toast.Show("❌ No text selected", 3000)
    return
  }

  ; The text is already selected from the ^c operation above
  ; We don't need to reselect it, just start proofreading
  ; The selection will be replaced directly

  ; Start proofreading
  ProofreadText(selectedText, geminiAPIkey)
}

; View log - Ctrl+Shift+L
!l:: {
  if (FileExist(logFile)) {
    Run("notepad.exe " . logFile)
  } else {
    Toast.Show("📄 No log file yet", 3000)
  }
}

; Removed Esc tooltip-hiding hotkey (no tooltips used)

; Initialize
InitializeLog()

; Simple startup message
Toast.Show("📝 ProofixAI is ready!`nSelect text and press Alt+P", 5000)