import '@k2oss/k2-broker-core';
import './index.ts';

function mock(name: string, value: any) {
    global[name] = value;
}

// This value is obfuscated on purpose. Replace with a valid OAuth token to run
let OAuthToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IjU3UDIzTms5bWQwVVpaVzJKNGFxS2xJcEpGMDR0YjhUdFg5bzJtTjFEZ1EiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mMGNkM2U0ZC00NDAyLTQ5NDAtODQyYi03ZjRjZGU2YjRjM2YvIiwiaWF0IjoxNjIxMjU0Mjk3LCJuYmYiOjE2MjEyNTQyOTcsImV4cCI6MTYyMTM0MDk5NiwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iLCJ1cm46bWljcm9zb2Z0OnJlcTEiLCJ1cm46bWljcm9zb2Z0OnJlcTIiLCJ1cm46bWljcm9zb2Z0OnJlcTMiLCJjMSIsImMyIiwiYzMiLCJjNCIsImM1IiwiYzYiLCJjNyIsImM4IiwiYzkiLCJjMTAiLCJjMTEiLCJjMTIiLCJjMTMiLCJjMTQiLCJjMTUiLCJjMTYiLCJjMTciLCJjMTgiLCJjMTkiLCJjMjAiLCJjMjEiLCJjMjIiLCJjMjMiLCJjMjQiLCJjMjUiXSwiYWlvIjoiQVNRQTIvOFRBQUFBSUpNck1hWVhudlJ1UEpQNjl1VkxGeXlpNEtLZ2RrZE9xTTNQQ244bTdvdz0iLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIGV4cGxvcmVyIChvZmZpY2lhbCBzaXRlKSIsImFwcGlkIjoiZGU4YmM4YjUtZDlmOS00OGIxLWE4YWQtYjc0OGRhNzI1MDY0IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJIYXl0ZXIiLCJnaXZlbl9uYW1lIjoiRXJuaWUiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiI4Mi4xMi4xNTYuMTg0IiwibmFtZSI6IkVybmllIEhheXRlciIsIm9pZCI6IjlkMzgyNGI1LThkNDItNDU4YS04Mjc1LTdhNjAwMzdhZDkyNSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMjAwMEEyMjY0NDREIiwicmgiOiIwLkFSQUFUVDdOOEFKRVFFbUVLMzlNM210TVA3WElpOTc1MmJGSXFLMjNTTnB5VUdRUUFKSS4iLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENoYW5uZWxNZW1iZXIuUmVhZC5BbGwgQ2hhbm5lbE1lbWJlci5SZWFkV3JpdGUuQWxsIENoYW5uZWxNZXNzYWdlLlJlYWQuQWxsIENoYW5uZWxNZXNzYWdlLlNlbmQgQ29udGFjdHMuUmVhZFdyaXRlIERldmljZU1hbmFnZW1lbnRBcHBzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUHJpdmlsZWdlZE9wZXJhdGlvbnMuQWxsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50UkJBQy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50UkJBQy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRTZXJ2aWNlQ29uZmlnLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRTZXJ2aWNlQ29uZmlnLlJlYWRXcml0ZS5BbGwgRGlyZWN0b3J5LkFjY2Vzc0FzVXNlci5BbGwgRGlyZWN0b3J5LlJlYWRXcml0ZS5BbGwgRmlsZXMuUmVhZFdyaXRlLkFsbCBHcm91cC5SZWFkLkFsbCBHcm91cC5SZWFkV3JpdGUuQWxsIElkZW50aXR5Umlza0V2ZW50LlJlYWQuQWxsIE1haWwuUmVhZFdyaXRlIE1haWxib3hTZXR0aW5ncy5SZWFkV3JpdGUgTm90ZXMuUmVhZFdyaXRlLkFsbCBvcGVuaWQgUGVvcGxlLlJlYWQgUG9saWN5LlJlYWRXcml0ZS5BcHBsaWNhdGlvbkNvbmZpZ3VyYXRpb24gcHJvZmlsZSBSZXBvcnRzLlJlYWQuQWxsIFNpdGVzLlJlYWQuQWxsIFNpdGVzLlJlYWRXcml0ZS5BbGwgVGFza3MuUmVhZFdyaXRlIFRlYW1zVGFiLlJlYWRXcml0ZS5BbGwgVGVhbXNUYWIuUmVhZFdyaXRlRm9yVGVhbSBVc2VyLlJlYWQgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6InhZdjU3T1RuaTdsRWs0elpSWnFlM01rdWxJd0tQLVRuLXJpelVSTnBLVXciLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiRVUiLCJ0aWQiOiJmMGNkM2U0ZC00NDAyLTQ5NDAtODQyYi03ZjRjZGU2YjRjM2YiLCJ1bmlxdWVfbmFtZSI6ImVybmllQGRlbmFsbGl4dGVjaGV1Lm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImVybmllQGRlbmFsbGl4dGVjaGV1Lm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IklmZFZFVFJyQzBPN1VUS2JrSmRDQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoiOTBhOE5fUTVEQUhOTmVvckt4WEcxZnBSQ2VvWWxsMVR3c0o1UVNPZ3ZGQSJ9LCJ4bXNfdGNkdCI6MTQzMTYzMDMwMX0.FIAMdlURYxpj4ktuk7OiOiOUOLr7BoJ2VnX1o_u_sxo9LfFR_tx_P5CCRAhRwNlwqH1d-Zpv_o3P_Z67l9wyFqk_Lvi9FmTMNv7GVUA8Fr_e2YZbHJD24FQcgKk9_jdVmGGGZy6wO0CmvxR8p5E-tikPnMou4Dc0ndlqKKdnlAyi2kw4uFvchmwqz99_kzNSxztSJxIHgCUjti0nniMz34AaQGZYuVzg6siWJhPUEmYp3ZJVDsaUdEnnUL6GlSZuSEv0YP9-Q2WXGWd7kryyIqYE2552dhab59yn_pXfxnMBZZV6SkuaGJGYyQyX05L_oNt9gPod2nVAajmUGxKJTw";

let schema = null;
mock('postSchema', function (result: any) {
    schema = result;
    console.log("postSchema:");
    console.log(schema);
});

let result: any = null;
function pr(r: any) {
    result = r;
   // console.log("postResult:")
   // console.log(JSON.stringify(result));
}

mock('postResult', pr);
let xhr: { [key: string]: any } = null;
class XHR {
    public onreadystatechange: () => void;
    public readyState: number;
    public status: number;
    public responseText: string;
    public withCredentials: boolean

    private recorder: { [key: string]: any };

    constructor() {
        xhr = this.recorder = {};
        this.recorder.headers = {};
    }

    open(method: string, url: string) {
        this.recorder.opened = { method, url };
    }

    setRequestHeader(key: string, value: string) {
        this.recorder.headers[key] = value;
       // console.log("setRequestHeader: " + key + "=" + value);
    }

    send(payload) {
        const request = require('request')
        if (this.withCredentials) {
            this.setRequestHeader("Authorization", "Bearer " + OAuthToken);
        }

        const options = {
            method: this.recorder.opened.method,
            url: this.recorder.opened.url,
            headers: this.recorder.headers,
            body: payload,
            strictSSL: false
        };
       // console.log("URL: " + options.method + " " + options.url);
       // console.log("BODY: " + options.body);
        let promise = new Promise((resolve, reject) => {
            try {
                request(options, (error, res, body) => {
                    if (error) {
                        console.error("error inside request:" + error)
                        return
                    }
                    this.responseText = body;
                    this.readyState = 4;
                    this.status = res.statusCode;
                    this.onreadystatechange();
                    resolve(body);
                    delete this.responseText;
                });
            }
            catch (err) {
                console.log("error ouside request " + err);
                reject()
            }
        }).catch((errr) => {
            console.log("Promise error:" + errr);
        });
    }
}

mock('XMLHttpRequest', XHR);

onexecute({
    objectName: 'excel',
    methodName: 'getUsedRangeItems',
    properties: {"fileId": "0135B642TVVZFROTARNZGKDJ6Y6WKQTSCT", "sheetName":"Sheet1"},
    parameters: {},
    configuration: {},
    schema: {}
});

