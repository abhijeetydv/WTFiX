/*
WTFiX - What The Fix!

Author: abhijeetydv
Date: 17-August-2025

WHAT IT DOES:
- Reads an API key from geminiAPIKey.txt (must be created manually).
- Select text, press Alt+Q: Instantly fix grammar and spelling.

CREDITS:
cJson.ahk - High-performance JSON library for AutoHotkey
Author: G33kDude
Source: https://github.com/G33kDude/cJson.ahk

Gemini API integration code by u/Laser_Made
Author: https://www.reddit.com/user/Laser_Made/
Source: https://www.reddit.com/r/AutoHotkey/comments/1ci2x6q/comment/l2hrijw/
*/

#Requires AutoHotkey v2.0+
#SingleInstance Force
#include JSON.ahk

; --- Global Variable ---
; Reads the key from geminiAPIKey.txt, which must be in the same folder as the script.
geminiAPIkey := Trim(FileRead(A_ScriptDir . "\geminiAPIKey.txt", "UTF-8"))

; Check if the API key was loaded successfully
if (geminiAPIkey = "") {
    MsgBox("Error: geminiAPIKey.txt is missing or empty.`n`nPlease create the file in the same folder as the script and paste your API key inside it.`n`nThe script will now exit.")
    ExitApp
}

; =================================================================
; ========== SCRIPT LOGIC BEGINS BEYOND THIS POINT ================
; =================================================================

; --- Main function to proofread text via API ---
ProofreadText(inputText) {
    prompt := "Proofread the following text for grammar, spelling, and punctuation errors. Provide only the corrected text. Do not include any introductory or concluding remarks. Text: `"" . inputText . "`""
    strUrl := "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?key=" . geminiAPIkey
    
    ; Show a "Thinking..." status tooltip near the mouse cursor
    MouseGetPos(&mouseX, &mouseY)
    ToolTip("Thinking...", mouseX + 10, mouseY + 10)
    
    api := ComObject("MSXML2.XMLHTTP")
    
    try {
        api.Open("POST", strUrl, false)
        api.SetRequestHeader("Content-Type", "application/json")
        api.Send(JSON.Dump({contents:[{parts:[{text:prompt}]}]}))
        
        ; Wait for the API response
        while (api.readyState != 4) {
            Sleep(100)
        }
        
        ; Handle API errors
        if (api.status != 200) {
            MouseGetPos(&mouseX, &mouseY)
            ToolTip("API Error`nCheck your API key and billing status.", mouseX + 10, mouseY + 10)
            SetTimer(() => ToolTip(), -4000)
            return ""
        }
        
        ; Process the streamed response and replace the selected text
        response := api.responseText
        correctedText := ProcessAndReplace(response)
        
        ; Show a completion status tooltip
        MouseGetPos(&mouseX, &mouseY)
        if (correctedText != inputText) {
            ToolTip("Text corrected!", mouseX + 10, mouseY + 10)
        } else {
            ToolTip("No changes needed.", mouseX + 10, mouseY + 10)
        }
        SetTimer(() => ToolTip(), -1500)
        
        return correctedText
        
    } catch as e {
        MouseGetPos(&mouseX, &mouseY)
        ToolTip("Connection failed", mouseX + 10, mouseY + 10)
        SetTimer(() => ToolTip(), -2000)
        return ""
    }
}

; --- Process the streamed API response and replace text ---
ProcessAndReplace(response) {
    try {
        ; Clean up the raw response and split it into individual JSON objects
        cleanResponse := Trim(response, "[]`n`r `t")
        jsonObjects := SplitJSON(cleanResponse)
        
        ; Step 1: Delete the user's selected text
        Send("{Delete}")
        
        fullText := "" ; Variable to assemble the full corrected text
        
        ; Step 2: Assemble the full corrected text in the background
        for _, jsonStr in jsonObjects {
            try {
                data := JSON.load(jsonStr)
                if (data.Has("candidates") && data["candidates"].Length > 0) {
                    candidate := data["candidates"][1]
                    if (candidate.Has("content") && candidate["content"].Has("parts") && candidate["content"]["parts"].Length > 0) {
                        newChunk := candidate["content"]["parts"][1]["text"]
                        fullText .= newChunk ; Add chunk to the full string
                    }
                }
            } catch {
                continue ; Skip any malformed JSON chunks
            }
        }

        fullText := Trim(fullText) ; Final trim for safety

        ; Step 3: Paste the complete text via the clipboard
        oldClipboard := A_Clipboard
        A_Clipboard := fullText
        SendInput("^v")
        Sleep(200) ; Brief pause to ensure paste completes before restoring clipboard
        A_Clipboard := oldClipboard
        
        return fullText

    } catch {
        MouseGetPos(&mouseX, &mouseY)
        ToolTip("Processing failed", mouseX + 10, mouseY + 10)
        SetTimer(() => ToolTip(), -2000)
        return ""
    }
}

; --- Helper function to split streamed JSON chunks ---
SplitJSON(jsonString) {
    jsonObjects := []
    currentObject := ""
    braceCount := 0
    inString := false
    escapeNext := false
    
    Loop Parse, jsonString {
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

; --- Hotkey to trigger WTFiX: Alt+Q ---
!q:: {
    ; Copy selected text to a variable
    oldClipboard := A_Clipboard
    A_Clipboard := ""
    SendInput("^c")
    
    if (!ClipWait(1)) {
        MouseGetPos(&mouseX, &mouseY)
        ToolTip("No text selected", mouseX + 10, mouseY + 10)
        SetTimer(() => ToolTip(), -1500)
        return
    }
    
    selectedText := A_Clipboard
    A_Clipboard := oldClipboard ; Restore original clipboard
    
    if (selectedText == "" || StrLen(Trim(selectedText)) < 1) {
        MouseGetPos(&mouseX, &mouseY)
        ToolTip("No text selected", mouseX + 10, mouseY + 10)
        SetTimer(() => ToolTip(), -1500)
        return
    }
    
    ; Call the main function to process the selected text
    ProofreadText(selectedText)
}

; --- Hotkey to hide any active tooltip: Esc ---
Esc::ToolTip()

; --- Startup message ---
MouseGetPos(&mouseX, &mouseY)
ToolTip("WTFiX is ready!`nSelect text and press Alt+Q", mouseX + 10, mouseY + 10)
SetTimer(() => ToolTip(), -4000)