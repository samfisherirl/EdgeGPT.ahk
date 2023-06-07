#Requires AutoHotkey v2.0.2
#SingleInstance
#Include "_jxon.ahk"
Persistent

/*
====================================================
Script Tray Menu
====================================================
*/

ReloadScript(*) {
	Reload
}

Debug(*) {
	ListLines
}

Exit(*) {
	ExitApp
}

/*
====================================================
Dark mode menu
====================================================
*/

Class DarkMode {
    Static __New(Mode := 1) => ( ; Mode: Dark = 1, Default (Light) = 0
        DllCall(DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", 135, "ptr"), "int", mode),
        DllCall(DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", 136, "ptr"))
    )
}

/*
====================================================
Variables
====================================================
*/

API_URL := "https://edgeservices.bing.com/edgesvc/turing/conversation/create"
Status_Message := ""
Response_Window_Status := "Closed"
Retry_Status := ""

/*
====================================================
Menus and ChatGPT prompts
====================================================
*/

MenuPopup := Menu()
MenuPopup.Add("&1 - Rephrase", Rephrase)
MenuPopup.Add("&2 - Summarize", Summarize)
MenuPopup.Add("&3 - Explain", Explain)
MenuPopup.Add("&4 - Expand", Expand)
MenuPopup.Add()
MenuPopup.Add("&5 - Generate reply", GenerateReply)
MenuPopup.Add("&6 - Find action items", FindActionItems)
MenuPopup.Add("&7 - Translate to English", TranslateToEnglish)

Rephrase(*) {
    ChatGPT_Prompt := "Rephrase the following text or paragraph to ensure clarity, conciseness, and a natural flow. The revision should preserve the tone, style, and formatting of the original text. Additionally, correct any grammar and spelling errors you come across:"
    Status_Message := "Rephrasing..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

Summarize(*) {
    ChatGPT_Prompt := "Summarize the following:"
    Status_Message := "Summarizing..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

Explain(*) {
    ChatGPT_Prompt := "Explain the following:"
    Status_Message := "Explaining..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

Expand(*) {
    ChatGPT_Prompt := "Considering the original tone, style, and formatting, please help me express the following idea in a clearer and more articulate way. The style of the message could be formal, informal, casual, empathetic, assertive, or persuasive, depending on the context of the original message. The text should be divided into paragraphs for readability. No specific language complexities need to be avoided and the focus should be equally distributed throughout the message. There's no set minimum or maximum length. Here's what I'm trying to say:"
    Status_Message := "Expanding..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

GenerateReply(*) {
    ChatGPT_Prompt := "Craft a response to any given message. The response should adhere to the original sender's tone, style, formatting, and cultural or regional context. Maintain the same level of formality and emotional tone as the original message. Responses may be of any length, provided they effectively communicate the response to the original sender:"
    Status_Message := "Generating reply..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

FindActionItems(*) {
    ChatGPT_Prompt := "Find action items that needs to be done and present them in a list:"
    Status_Message := "Finding action items..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

TranslateToEnglish(*) {
    ChatGPT_Prompt := "Generate an English translation for the following text or paragraph, ensuring the translation accurately conveys the intended meaning or idea without excessive deviation. The translation should preserve the tone, style, and formatting of the original text:"
    Status_Message := "Translating to English..."
    ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status)
}

/*
====================================================
Create Response Window
====================================================
*/

Response_Window := Gui("-Caption", "Response")
Response_Window.BackColor := "0x333333"
Response_Window.SetFont("s13 cWhite", "Georgia")
Response := Response_Window.Add("Edit", "r20 ReadOnly w600 Wrap Background333333", Status_Message)
RetryButton := Response_Window.Add("Button", "x190 Disabled", "Retry")
RetryButton.OnEvent("Click", Retry)
CopyButton := Response_Window.Add("Button", "x+30 w80 Disabled", "Copy")
CopyButton.OnEvent("Click", Copy)
Response_Window.Add("Button", "x+30", "Close").OnEvent("Click", Close)

Response_Window.Show()
/*
====================================================
Buttons
====================================================
*/

Retry(*) {
    Retry_Status := "Retry"
    RetryButton.Enabled := 0
    CopyButton.Enabled := 0
    CopyButton.Text := "Copy"
    ProcessRequest(Previous_ChatGPT_Prompt, Previous_Status_Message, Retry_Status)
}

Copy(*) {
    A_Clipboard := Response.Value
    CopyButton.Enabled := 0
    CopyButton.Text := "Copied!"

    DllCall("SetFocus", "Ptr", 0)
    Sleep 2000

    CopyButton.Enabled := 1
    CopyButton.Text := "Copy"
}

Close(*) {
    HTTP_Request.Abort
    Response_Window.Hide
    global Response_Window_Status := "Closed"
}

/*
====================================================
Connect to ChatGPT API and process request
====================================================
*/

ProcessRequest(ChatGPT_Prompt, Status_Message, Retry_Status) {
    if (Retry_Status != "Retry") {
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2) {
            MsgBox "The attempt to copy text onto the clipboard failed."
            return
        }
        CopiedText := A_Clipboard
        ChatGPT_Prompt := ChatGPT_Prompt "`n`n" CopiedText
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, '(\\|")+', '\$1') ; Clean back spaces and quotes
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, "`n", "\n") ; Clean newlines
        ChatGPT_Prompt := RegExReplace(ChatGPT_Prompt, "`r", "") ; Remove carriage returns
        global Previous_ChatGPT_Prompt := ChatGPT_Prompt
        global Previous_Status_Message := Status_Message
        global Response_Window_Status
    }

    OnMessage 0x200, WM_MOUSEHOVER
    Response.Value := Status_Message
    if (Response_Window_Status = "Closed") {
        Response_Window.Show("AutoSize Center")
        Response_Window_Status := "Open"
        RetryButton.Enabled := 0
        CopyButton.Enabled := 0
    }    
    DllCall("SetFocus", "Ptr", 0)

    global HTTP_Request := ComObject("WinHttp.WinHttpRequest.5.1")
    HTTP_Request.open("GET", API_URL, true)
    HTTP_Request.SetRequestHeader("accept", "application/json")
    HTTP_Request.SetRequestHeader("accept-language", "en-US,en;q=0.9")
    HTTP_Request.SetRequestHeader("content-type", "application/json")
    HTTP_Request.SetRequestHeader("sec-ch-ua", '"Not_A Brand";v="99") "Microsoft Edge";v="110") "Chromium";v="110"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-arch", '"x86"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-bitness", '"64"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-full-version", '"109.0.1518.78"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-full-version-list", '"Chromium";v="110.0.5481.192") "Not A(Brand";v="24.0.0.0") "Microsoft Edge";v="110.0.1587.69"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-mobile", "?0")
    HTTP_Request.SetRequestHeader("sec-ch-ua-model", "")
    HTTP_Request.SetRequestHeader("sec-ch-ua-platform", '"Windows"')
    HTTP_Request.SetRequestHeader("sec-ch-ua-platform-version", '"15.0.0"')
    HTTP_Request.SetRequestHeader("sec-fetch-dest", "empty")
    HTTP_Request.SetRequestHeader("sec-fetch-mode", "cors")
    HTTP_Request.SetRequestHeader("sec-fetch-site", "same-origin")
    ; HTTP_Request.SetRequestHeader("x-ms-client-request-id", str(uuid.uuid4())
    HTTP_Request.SetRequestHeader("x-ms-useragent", "azsdk-js-api-client-factory/1.0.0-beta.1 core-rest-pipeline/1.10.0 OS/Win32")
    HTTP_Request.SetRequestHeader("Referer", "https://www.bing.com/search?q=Bing+AI&showconv=1&FORM=hpcodx")
    HTTP_Request.SetRequestHeader("Referrer-Policy", "origin-when-cross-origin")
    FORWARDED_IP := "13." Random(104, 107) "."  Random(0, 255) "." Random(0, 255)
    HTTP_Request.SetRequestHeader("x-forwarded-for", FORWARDED_IP)
    
    HTTP_Request.SetTimeouts(60000, 60000, 60000, 60000)
    HTTP_Request.Send()
    SetTimer LoadingCursor, 1
    if WinExist("Response") {
        WinActivate "Response"
    }
    HTTP_Request.WaitForResponse
    try {
        if (HTTP_Request.status == 200) {
            JSON_Response := HTTP_Request.responseText
            var := Jxon_Load(&JSON_Response)
            JSON_Response := var.Get("choices")[1].Get("message").Get("content")
            RetryButton.Enabled := 1
            CopyButton.Enabled := 1
            Response.Value := JSON_Response

            SetTimer LoadingCursor, 0
            OnMessage 0x200, WM_MOUSEHOVER, 0
            Cursor := DllCall("LoadCursor", "uint", 0, "uint", 32512) ; Arrow cursor
            DllCall("SetCursor", "UPtr", Cursor)

            Response_Window.Flash()
            DllCall("SetFocus", "Ptr", 0)
        } else {
            RetryButton.Enabled := 1
            CopyButton.Enabled := 1
            Response.Value := "Status " HTTP_Request.status " " HTTP_Request.responseText

            SetTimer LoadingCursor, 0
            OnMessage 0x200, WM_MOUSEHOVER, 0
            Cursor := DllCall("LoadCursor", "uint", 0, "uint", 32512) ; Arrow cursor
            DllCall("SetCursor", "UPtr", Cursor)

            Response_Window.Flash()
            DllCall("SetFocus", "Ptr", 0)
        }
    }
}

/*
====================================================
Cursors
====================================================
*/

WM_MOUSEHOVER(*) {
    Cursor := DllCall("LoadCursor", "uint", 0, "uint", 32648) ; Unavailable cursor
    MouseGetPos ,,, &MousePosition
    if (CopyButton.Enabled = 0) & (MousePosition = "Button2") {
        DllCall("SetCursor", "UPtr", Cursor)
    } else if (RetryButton.Enabled = 0) & (MousePosition = "Button1") | (MousePosition = "Button2") {
        DllCall("SetCursor", "UPtr", Cursor)
    }
}

LoadingCursor() {
    MouseGetPos ,,, &MousePosition
    if (MousePosition = "Edit1") {
        Cursor := DllCall("LoadCursor", "uint", 0, "uint", 32514) ; Loading cursor
        DllCall("SetCursor", "UPtr", Cursor)
    }
}

/*
====================================================
Hotkey
====================================================
*/

^::MenuPopup.Show()