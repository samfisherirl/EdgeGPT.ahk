API_URL := "https://edgeservices.bing.com/edgesvc/turing/conversation/create"


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
HTTP_Request.SetRequestHeader("sec-ch-ua-platform", '"Windows"')
HTTP_Request.SetRequestHeader("sec-ch-ua-platform-version", '"15.0.0"')
HTTP_Request.SetRequestHeader("sec-fetch-dest", "empty")
HTTP_Request.SetRequestHeader("sec-fetch-mode", "cors")
HTTP_Request.SetRequestHeader("sec-fetch-site", "same-origin")
; HTTP_Request.SetRequestHeader("x-ms-client-request-id", str(uuid.uuid4())
HTTP_Request.SetRequestHeader("x-ms-useragent", "azsdk-js-api-client-factory/1.0.0-beta.1 core-rest-pipeline/1.10.0 OS/Win32")
HTTP_Request.SetRequestHeader("Referer", "https://www.bing.com/search?q=Bing+AI&showconv=1&FORM=hpcodx")
HTTP_Request.SetRequestHeader("Referrer-Policy", "origin-when-cross-origin")
FORWARDED_IP := "13." Random(104, 107) "." Random(0, 255) "." Random(0, 255)
HTTP_Request.SetRequestHeader("x-forwarded-for", FORWARDED_IP)

HTTP_Request.SetTimeouts(60000, 60000, 60000, 60000)
HTTP_Request.Send()
if WinExist("Response") {
    WinActivate "Response"
}
HTTP_Request.WaitForResponse
msg := HTTP_Request.ResponseText
MsgBox(msg)