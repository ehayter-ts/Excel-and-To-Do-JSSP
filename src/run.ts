import '@k2oss/k2-broker-core';
import './index.ts';

function mock(name: string, value: any) {
    global[name] = value;
}

// This value is obfuscated on purpose. Replace with a valid OAuth token to run
let OAuthToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6InNsTkNXR0pMR0pMRW5yLTRnYTdXT2tJd0NfWWo0UU5CQkotTzFJM1lQZm8iLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mMGNkM2U0ZC00NDAyLTQ5NDAtODQyYi03ZjRjZGU2YjRjM2YvIiwiaWF0IjoxNjIwOTA5NzI2LCJuYmYiOjE2MjA5MDk3MjYsImV4cCI6MTYyMDk5NjQyNSwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iLCJ1cm46bWljcm9zb2Z0OnJlcTEiLCJ1cm46bWljcm9zb2Z0OnJlcTIiLCJ1cm46bWljcm9zb2Z0OnJlcTMiLCJjMSIsImMyIiwiYzMiLCJjNCIsImM1IiwiYzYiLCJjNyIsImM4IiwiYzkiLCJjMTAiLCJjMTEiLCJjMTIiLCJjMTMiLCJjMTQiLCJjMTUiLCJjMTYiLCJjMTciLCJjMTgiLCJjMTkiLCJjMjAiLCJjMjEiLCJjMjIiLCJjMjMiLCJjMjQiLCJjMjUiXSwiYWlvIjoiRTJaZ1lOZ1I5ZjVZZVJybko2M2dxdXh2LzIyVEJGMUZORmhNL05tV1ZUVkpMallQV1FBQSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggZXhwbG9yZXIgKG9mZmljaWFsIHNpdGUpIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkhheXRlciIsImdpdmVuX25hbWUiOiJFcm5pZSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjgyLjEyLjE1Ni4xODQiLCJuYW1lIjoiRXJuaWUgSGF5dGVyIiwib2lkIjoiOWQzODI0YjUtOGQ0Mi00NThhLTgyNzUtN2E2MDAzN2FkOTI1IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAwQTIyNjQ0NEQiLCJyaCI6IjAuQVJBQVRUN044QUpFUUVtRUszOU0zbXRNUDdYSWk5NzUyYkZJcUsyM1NOcHlVR1FRQUpJLiIsInNjcCI6IkNhbGVuZGFycy5SZWFkV3JpdGUgQ2hhbm5lbE1lbWJlci5SZWFkLkFsbCBDaGFubmVsTWVtYmVyLlJlYWRXcml0ZS5BbGwgQ2hhbm5lbE1lc3NhZ2UuUmVhZC5BbGwgQ2hhbm5lbE1lc3NhZ2UuU2VuZCBDb250YWN0cy5SZWFkV3JpdGUgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50Q29uZmlndXJhdGlvbi5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50Q29uZmlndXJhdGlvbi5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5Qcml2aWxlZ2VkT3BlcmF0aW9ucy5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRSQkFDLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRSQkFDLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFNlcnZpY2VDb25maWcuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFNlcnZpY2VDb25maWcuUmVhZFdyaXRlLkFsbCBEaXJlY3RvcnkuQWNjZXNzQXNVc2VyLkFsbCBEaXJlY3RvcnkuUmVhZFdyaXRlLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgSWRlbnRpdHlSaXNrRXZlbnQuUmVhZC5BbGwgTWFpbC5SZWFkV3JpdGUgTWFpbGJveFNldHRpbmdzLlJlYWRXcml0ZSBOb3Rlcy5SZWFkV3JpdGUuQWxsIG9wZW5pZCBQZW9wbGUuUmVhZCBQb2xpY3kuUmVhZFdyaXRlLkFwcGxpY2F0aW9uQ29uZmlndXJhdGlvbiBwcm9maWxlIFJlcG9ydHMuUmVhZC5BbGwgU2l0ZXMuUmVhZC5BbGwgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUYXNrcy5SZWFkV3JpdGUgVGVhbXNUYWIuUmVhZFdyaXRlLkFsbCBUZWFtc1RhYi5SZWFkV3JpdGVGb3JUZWFtIFVzZXIuUmVhZCBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIGVtYWlsIiwic2lnbmluX3N0YXRlIjpbImttc2kiXSwic3ViIjoieFl2NTdPVG5pN2xFazR6WlJacWUzTWt1bEl3S1AtVG4tcml6VVJOcEtVdyIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJFVSIsInRpZCI6ImYwY2QzZTRkLTQ0MDItNDk0MC04NDJiLTdmNGNkZTZiNGMzZiIsInVuaXF1ZV9uYW1lIjoiZXJuaWVAZGVuYWxsaXh0ZWNoZXUub25taWNyb3NvZnQuY29tIiwidXBuIjoiZXJuaWVAZGVuYWxsaXh0ZWNoZXUub25taWNyb3NvZnQuY29tIiwidXRpIjoiQTdyTzliWVBPa0sycU9INjhBVklBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19zdCI6eyJzdWIiOiI5MGE4Tl9RNURBSE5OZW9yS3hYRzFmcFJDZW9ZbGwxVHdzSjVRU09ndkZBIn0sInhtc190Y2R0IjoxNDMxNjMwMzAxfQ.ln3gHGpgO0T1KqGXeSp9WokS-Ai4J1iL4iiG5t9nh9bIGIM7-dUnYAcah26CJGOEqjryuC_TGE6BI4VtFMf4l8CeAEweSmIeLnjiBZO8dTnQDUjRs7r4g_8iI9qo4vIcIYH-1TiJ-0HCE2CRHkcehaVwy4mZYHcBfJ1lQcUapsWl4-y7y-guMa2fLkAMOFVmEjNNxw7qxkyIT0iilZuypVcABBvpg2ATly5inNXDVMg-AloYN3aqpLTMmmaJwJcodmxZlGmn9OS76qCMhD6UvsH04sQjzP0BpHb-NIS6tOng85pQqs6To0uaqmHhwFg7o5IyWAVpRqgmcYI17EDcLg";

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
    objectName: 'drive',
    methodName: 'createFolder',
    properties: {"folderName": "Ernie Test"},
    parameters: {},
    configuration: {},
    schema: {}
});

