#NoEnv

; Import required COM libraries
ComObjError(false)
WinHTTP := ComObjCreate("WinHTTP.WinHTTPRequest.5.1")

class Chatbot {
    ; Combines everything to make it seamless
    __New(proxy := "", cookies := "") {
        this.proxy := proxy
        this.chat_hub := new ChatHub(this.proxy, cookies)
    }

    static async create(proxy := "", cookies := "") {
        self := new Chatbot(proxy, cookies)
        self.chat_hub := new ChatHub(await Conversation.create(self.proxy, cookies), self.proxy, cookies)
        return self
    }

    async save_conversation(filename) {
        conversation_id := this.chat_hub.request.conversation_id
        conversation_signature := this.chat_hub.request.conversation_signature
        client_id := this.chat_hub.request.client_id
        invocation_id := this.chat_hub.request.invocation_id

        jsonObj := {
            "conversation_id": conversation_id,
            "conversation_signature": conversation_signature,
            "client_id": client_id,
            "invocation_id": invocation_id
        }

        File := FileOpen(filename, "w")
        File.Write(jsonObj)
        File.Close()
    }

    async load_conversation(filename) {
        File := FileOpen(filename, "r")
        jsonStr := File.Read()
        File.Close()

        jsonObj := Json.Load(jsonStr)

        this.chat_hub.request := new ChatHubRequest(
            conversation_signature := jsonObj["conversation_signature"],
            client_id := jsonObj["client_id"],
            conversation_id := jsonObj["conversation_id"],
            invocation_id := jsonObj["invocation_id"]
        )
    }

    async get_conversation() {
        return await this.chat_hub.get_conversation()
    }

    async ask(prompt, wss_link := "wss://sydney.bing.com/sydney/ChatHub", conversation_style := "", webpage_context := "", search_result := false, locale := "") {
        loop {
            for response in this.chat_hub.ask_stream(prompt, conversation_style, wss_link, webpage_context, search_result, locale) {
                final := response[0]
                responseData := response[1]

                if final {
                    return responseData
                }
            }
            this.chat_hub.wss.close()
        }
    }

    async ask_stream(prompt, wss_link := "wss://sydney.bing.com/sydney/ChatHub", conversation_style := "", raw := false, webpage_context := "", search_result := false, locale := "") {
        for response in this.chat_hub.ask_stream(prompt, conversation_style, wss_link, raw, webpage_context, search_result, locale) {
            yield response
        }
    }

    async close() {
        await this.chat_hub.close()
    }

    async reset() {
        await this.close()
        this.chat_hub := new ChatHub(await Conversation.create(this.proxy, this.chat_hub.cookies), this.proxy, this.chat_hub.cookies)
    }
}

ChatHub := class {
    __New(conversation := "", proxy := "", cookies := "") {
        this.request := new Request(conversation, proxy, cookies)
        this.proxy := proxy
        this.cookies := cookies
    }

    async get_conversation() {
        return await this.request.get_conversation()
    }

    async ask_stream(prompt, conversation_style := "", wss_link := "wss://sydney.bing.com/sydney/ChatHub", raw := false, webpage_context := "", search_result := false, locale := "") {
        return await this.request.ask_stream(prompt, conversation_style, wss_link, raw, webpage_context, search_result, locale)
    }

    async close() {
        await this.request.close()
    }
}

Request := class {
    __New(conversation := "", proxy := "", cookies := "") {
        this.conversation := conversation
        this.proxy := proxy
        this.cookies := cookies
    }

    async get_conversation() {
        url := "https://sydney.bing.com/sydney/conversations/%conversation%/activities?watermark=-1"
        url := StrReplace(url, "%conversation%", this.conversation)
        headers := ""

        if (this.cookies != "") {
            headers := "Cookie: " . this.cookies
        }

        WinHTTP.Open("GET", url, true)
        WinHTTP.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        WinHTTP.SetRequestHeader("Accept", "application/json")
        WinHTTP.SetRequestHeader("Referer", "https://www.bing.com/")
        WinHTTP.SetRequestHeader("Accept-Language", "en-US,en;q=0.9")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Extensions", "permessage-deflate")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Version", "13")
        WinHTTP.SetRequestHeader("Upgrade", "websocket")

        if (headers != "") {
            WinHTTP.SetRequestHeader("Cookie", headers)
        }

        WinHTTP.Send()
        WinHTTP.WaitForResponse()
        response := WinHTTP.ResponseText

        return Json.Load(response)
    }

    async ask_stream(prompt, conversation_style := "", wss_link := "wss://sydney.bing.com/sydney/ChatHub", raw := false, webpage_context := "", search_result := false, locale := "") {
        url := "https://sydney.bing.com/sydney/conversations"
        headers := ""

        if (this.cookies != "") {
            headers := "Cookie: " . this.cookies
        }

        body := {
            "query": prompt,
            "conversationId": this.conversation,
            "style": conversation_style,
            "watermark": -1,
            "type": "message",
            "messageId": -1,
            "payload": {
                "userMessageId": -1,
                "type": "message",
                "content": prompt,
                "searchResultPayload": {
                    "triggerSearch": search_result
                },
                "contentRequestPayload": {
                    "shouldAppendCurrentUserMessage": true,
                    "type": "RteDocument",
                    "version": 1,
                    "content": {
                        "type": "Node",
                        "id": "0",
                        "childs": [
                            {
                                "type": "Node",
                                "id": "1",
                                "childs": [
                                    {
                                        "type": "Node",
                                        "id": "2",
                                        "childs": [
                                            {
                                                "type": "Node",
                                                "id": "3",
                                                "childs": [],
                                                "props": {
                                                    "href": "",
                                                    "target": "",
                                                    "rel": "",
                                                    "onclick": "",
                                                    "onkeypress": "",
                                                    "tabIndex": -1,
                                                    "colSpan": 0,
                                                    "rowSpan": 0,
                                                    "data-ms-equation": "",
                                                    "style": ""
                                                },
                                                "name": "a"
                                            }
                                        ],
                                        "props": {
                                            "href": "",
                                            "target": "",
                                            "rel": "",
                                            "onclick": "",
                                            "onkeypress": "",
                                            "tabIndex": -1,
                                            "colSpan": 0,
                                            "rowSpan": 0,
                                            "data-ms-equation": "",
                                            "style": ""
                                        },
                                        "name": "a"
                                    }
                                ],
                                "props": {
                                    "href": "",
                                    "target": "",
                                    "rel": "",
                                    "onclick": "",
                                    "onkeypress": "",
                                    "tabIndex": -1,
                                    "colSpan": 0,
                                    "rowSpan": 0,
                                    "data-ms-equation": "",
                                    "style": ""
                                },
                                "name": "a"
                            }
                        ],
                        "props": {
                            "href": "",
                            "target": "",
                            "rel": "",
                            "onclick": "",
                            "onkeypress": "",
                            "tabIndex": -1,
                            "colSpan": 0,
                            "rowSpan": 0,
                            "data-ms-equation": "",
                            "style": ""
                        },
                        "name": "a"
                    },
                    "contentV2": {},
                    "contentType": "RteDocument"
                }
            }
        }

        if (webpage_context != "") {
            body["payload"]["webpageContext"] := webpage_context
        }

        if (locale != "") {
            body["locale"] := locale
        }

        WinHTTP.Open("POST", url, true)
        WinHTTP.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        WinHTTP.SetRequestHeader("Accept", "application/json")
        WinHTTP.SetRequestHeader("Referer", "https://www.bing.com/")
        WinHTTP.SetRequestHeader("Accept-Language", "en-US,en;q=0.9")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Extensions", "permessage-deflate")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Version", "13")
        WinHTTP.SetRequestHeader("Upgrade", "websocket")

        if (headers != "") {
            WinHTTP.SetRequestHeader("Cookie", headers)
        }

        WinHTTP.Send(Json.Dump(body))
        WinHTTP.WaitForResponse()
        response := WinHTTP.ResponseText

        if (raw) {
            return response
        }

        return Json.Load(response)
    }
}

ChatHub := class {
    __New(conversation, proxy := "", cookies := "") {
        this.request := Request(conversation, proxy, cookies)
        this.conversation := conversation
        this.proxy := proxy
        this.cookies := cookies
        this.wss := ""
        this.close := false
    }

    async get_conversation() {
        return this.request.get_conversation()
    }

    async ask_stream(prompt, conversation_style := "", wss_link := "wss://sydney.bing.com/sydney/ChatHub", raw := false, webpage_context := "", search_result := false, locale := "") {
        if (!this.wss) {
            this.wss := ComObjCreate("WinHttp.WinHttpRequest.5.1")
            url := "https://sydney.bing.com/sydney/ChatHub"
            headers := ""

            if (this.cookies != "") {
                headers := "Cookie: " . this.cookies
            }

            this.wss.Open("GET", url, true)
            this.wss.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            this.wss.SetRequestHeader("Accept", "application/json")
            this.wss.SetRequestHeader("Referer", "https://www.bing.com/")
            this.wss.SetRequestHeader("Accept-Language", "en-US,en;q=0.9")
            this.wss.SetRequestHeader("Sec-WebSocket-Extensions", "permessage-deflate")
            this.wss.SetRequestHeader("Sec-WebSocket-Version", "13")
            this.wss.SetRequestHeader("Upgrade", "websocket")

            if (headers != "") {
                this.wss.SetRequestHeader("Cookie", headers)
            }

            this.wss.Send()
            this.wss.WaitForResponse()

            if (this.wss.Status == 101) {
                conversation_id := this.request.conversation
                client_id := this.request.conversation
                invocation_id := this.request.conversation

                this.wss_url := wss_link
                this.wss_url := StrReplace(this.wss_url, "%conversation_id%", conversation_id)
                this.wss_url := StrReplace(this.wss_url, "%client_id%", client_id)
                this.wss_url := StrReplace(this.wss_url, "%invocation_id%", invocation_id)

                this.wss_url := this.wss_url . "&client-id=" . client_id

                this.wss.Open("GET", this.wss_url, true)
                this.wss.SetRequestHeader("Sec-WebSocket-Extensions", "permessage-deflate")
                this.wss.SetRequestHeader("Sec-WebSocket-Version", "13")
                this.wss.SetRequestHeader("Upgrade", "websocket")
                this.wss.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                this.wss.Send()

                return
            }
        }

        response := this.request.ask_stream(prompt, conversation_style, this.wss_url, raw, webpage_context, search_result, locale)

        for value in response {
            if (this.close) {
                this.wss.close()
                break
            }

            final := value[0]
            responseData := value[1]

            yield [final, responseData]
        }
    }

    async close() {
        this.close := true
    }
}

Conversation := class {
    __New(proxy := "", cookies := "") {
        this.proxy := proxy
        this.cookies := cookies
    }

    async create(proxy := "", cookies := "") {
        conversation_id := ""

        url := "https://sydney.bing.com/sydney/conversations"
        headers := ""

        if (this.cookies != "") {
            headers := "Cookie: " . this.cookies
        }

        body := {
            "customData": {},
            "skillset": "custom",
            "styleOverride": "webchat"
        }

        if (proxy != "") {
            body["customData"]["com.microsoft.bot.builder.Proxy"] := proxy
        }

        WinHTTP.Open("POST", url, true)
        WinHTTP.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        WinHTTP.SetRequestHeader("Accept", "application/json")
        WinHTTP.SetRequestHeader("Referer", "https://www.bing.com/")
        WinHTTP.SetRequestHeader("Accept-Language", "en-US,en;q=0.9")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Extensions", "permessage-deflate")
        WinHTTP.SetRequestHeader("Sec-WebSocket-Version", "13")
        WinHTTP.SetRequestHeader("Upgrade", "websocket")

        if (headers != "") {
            WinHTTP.SetRequestHeader("Cookie", headers)
        }

        WinHTTP.Send(Json.Dump(body))
        WinHTTP.WaitForResponse()
        response := WinHTTP.ResponseText

        if (response) {
            jsonObj := Json.Load(response)
            conversation_id := jsonObj["conversationId"]
        }

        return conversation_id
    }
}

Json := {
    Load(str) {
        return JSON.parse(str)
    }

    Dump(obj) {
        return JSON.stringify(obj)
    }
}

; Usage example:
bot := Chatbot.create()
bot.ask("Hello, how are you?")