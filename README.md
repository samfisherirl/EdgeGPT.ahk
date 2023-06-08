# EdgeGPT.ahk

Working on converting EdgeGPT.py and BingChatAPI.cs to a pure autohotkey library. 
These libraries utilize Bing's public access with no cost or key required, utilizing ChatGPT-4

C#:  https://github.com/liaosunny123/BingChatApi/blob/master/BingChatApiLibs/BingChatClient.cs
Py: https://github.com/acheong08/EdgeGPT

### convoID grab:
```python
            response = await client.get(
                url=os.environ.get("BING_PROXY_URL")
                or "https://edgeservices.bing.com/edgesvc/turing/conversation/create",
            )
            if response.status_code != 200:
                response = await client.get(
                    "https://edge.churchless.tech/edgesvc/turing/conversation/create",
                )

```
### final post request
```python

                response = await client.post(
                    "https://sydney.bing.com/sydney/UpdateConversation/",
                    json={
                        "messages": [
                            {
                                "author": "user",
                                "description": webpage_context,
                                "contextType": "WebPage",
                                "messageType": "Context",
                            },
                        ],
                        "conversationId": self.request.conversation_id,
                        "source": "cib",
                        "traceId": _get_ran_hex(32),
                        "participant": {"id": self.request.client_id},
                        "conversationSignature": self.request.conversation_signature,
                    },
                )
            if response.status_code != 200:
                print(f"Status code: {response.status_code}")
                print(response.text)
                print(response.url)
```

struct needs: 

```python
        self.request = _ChatHubRequest(
            conversation_signature=conversation.struct["conversationSignature"],
            client_id=conversation.struct["clientId"],
            conversation_id=conversation.struct["conversationId"],
        )
```

conversation_id return

![image](https://github.com/samfisherirl/EdgeGPT.ahk/assets/98753696/3a3bdc02-2135-4ef9-950d-bce13c2fff65)

