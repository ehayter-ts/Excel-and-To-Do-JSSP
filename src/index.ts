import '@k2oss/k2-broker-core';

metadata = {
    systemName: "MSExcelTodoJSSP",
    displayName: "Microsoft Excel and To Do",
    description: "A connector for Microsoft Excel and To Do"
};

// Constants
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";
const baseUriEndpointBeta = "https://graph.microsoft.com/beta";

//
// Objects
const Drive = "drive";

//
//Drive
const FileId = "fileId";
const FileWebUrl = "fileWebUrl";
const FileSize = "fileSize";
const FileName = "fileName";
const FileCreated = "fileCreatedDate";
const FileCreatedBy = "fileCreatedBy";
const FileCreatedByEmail = "fileCreatedByEmail";
const FileMimeType = "fileMimeType";
const FilePath = "filePath";

const FileSearch = "searchFile";

//OnDescribe
ondescribe = function () {
    postSchema({

        objects: {
            [Drive]: {
                displayName: "Drive",
                description: "Drive",
                properties: {
                    [FileId]: {
                        displayName: "File ID",
                        description: "File ID",
                        type: "string"
                    },
                    [FileWebUrl]: {
                        displayName: "File Web URL",
                        description: "File Web URL",
                        type: "string"
                    },
                    [FileSize]: {
                        displayName: "File Size",
                        description: "File Size",
                        type: "number"
                    },
                    [FileName]: {
                        displayName: "File Name",
                        description: "File Name",
                        type: "string"
                    },
                    [FileCreated]: {
                        displayName: "File Created Date",
                        description: "File Created Date",
                        type: "dateTime"
                    },
                    [FileCreatedBy]: {
                        displayName: "File Created By",
                        description: "File Created By",
                        type: "string"
                    },
                    [FileCreatedByEmail]: {
                        displayName: "File Created By Email",
                        description: "File Created By Email",
                        type: "string"
                    },
                    [FileMimeType]: {
                        displayName: "File Mime Type",
                        description: "File Mime Type",
                        type: "string"
                    },
                    [FilePath]: {
                        displayName: "File Path",
                        description: "File Path",
                        type: "string"
                    }
                },
                methods: {
                    [FileSearch]: {
                        displayName: "Search for a file in my OneDrive",
                        type: "read",
                        inputs: [FileName, FilePath],
                        requiredInputs: [FileName],
                        outputs: [FileId, FileWebUrl, FileSize, FileName, FileCreated, FileCreatedBy, FileCreatedByEmail, FileMimeType]
                    }
                }
            }
        }

    });
}

// OnExecute
onexecute = function ({ objectName, methodName, parameters, properties }) {
    switch (objectName) {
        case Drive:
            onexecuteDrive(methodName, parameters, properties);
            break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteDrive(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case FileSearch:
            onexecuteSearchFile(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function ExecuteRequest(url: string, data: string, requestType: string, cb) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        console.log("ExecuteRequest XHR status: " + xhr.status + "," + xhr.responseText);
        if (xhr.readyState !== 4)
            return;
        if (xhr.status == 201) {
            //console.log("ExecuteRequest XHR status: " + xhr.status + "," + xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 204) {
            if (typeof cb === 'function')
                cb(xhr.responseText);
        }
        else if (xhr.status == 200) {
            var obj = JSON.parse(xhr.responseText);
            //console.log("ExecuteRequest XHR status: " + xhr.status + "," + xhr.responseText);
            //console.log("ExecuteRequest cb type of: " + (typeof cb).toString());
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 202) {
            if (typeof cb === 'function')
                cb(null);
        }
        else if (xhr.status == 400) {
            // This is a bad request, return error to UI
            var obj = JSON.parse(xhr.responseText);
            throw new Error(obj.error.code + ": " + obj.error.message);
        }
        else if (xhr.status == 404) {
            var obj = JSON.parse(xhr.responseText);
            // This is to supress an error that happens with team archive/unarchive
            var errorMessage = obj.error.message;
            if (errorMessage.startswith == "No Team found with Group id") {
                // do nothing - supress error
            }
            else {
                throw new Error(obj.error.code + ": " + obj.error.message);
            }
            //console.log("MSTeamsConnector ExecuteRequest: Failed with 404 error.");
            //throw new Error(obj.error.code + " error: " + obj.error.message);
            //console.log("Failed with status " + xhr.status + " " + xhr.responseText);
        }
        else {
            postResult({
                //TeamIsSuccessful: false
            });
            var obj = JSON.parse(xhr.responseText);
            throw new Error(obj.error.code + ": " + obj.error.message);
            //console.log("Failed with status " + xhr.status + " " + xhr.responseText);

        }
    };
    console.log("MSTeamsConnector ExecuteRequest: " + url);
    xhr.open(requestType.toUpperCase(), url);
    // Authentication Header
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");
    if (requestType.toUpperCase() == "PUT" || requestType.toUpperCase() == "POST" || requestType.toUpperCase() == "PATCH") {
        xhr.setRequestHeader("Content-Type", "application/json");
    }
    xhr.send(data);
}


function onexecuteSearchFile(parameters: SingleRecord, properties: SingleRecord) {
    GetDriveFile(parameters, properties, function (a) {
        var resultObj = {
            [FileId]: "", 
            [FileWebUrl]: "",
            [FileSize]: 0, 
            [FileName]: "", 
            [FileCreated]: null, 
            [FileCreatedBy]: "", 
            [FileCreatedByEmail]: "", 
            [FileMimeType]: ""
        };

        if (a.value.length > 0)
        {
            resultObj[FileId] = a.value[0].id; 
            resultObj[FileWebUrl] = a.value[0].webUrl;
            resultObj[FileSize] = a.value[0].size;
            resultObj[FileName] = a.value[0].name;
            resultObj[FileCreated] = a.value[0].createdDateTime;
            resultObj[FileCreatedBy] = a.value[0].createdBy.displayName;
            resultObj[FileCreatedByEmail] = a.value[0].createdBy.email;
            resultObj[FileMimeType] = a.value[0].file.mimeType;
        }

        postResult(resultObj);
    });
}

function GetDriveFile(parameters: SingleRecord, properties: SingleRecord, cb) {
    let fileName = properties[FileName];
    let filePath = properties[FilePath];
    if (!(typeof fileName === "string")) throw new Error("properties[FileName] is not of type string");

    var url;
    if ((typeof filePath === "string") && (filePath != ""))
    {
        url = baseUriEndpoint + `/me/drive/root:/${filePath}:/search(q='${fileName}')`;
    }
    else
    {
        url = baseUriEndpoint + `/me/drive/root/search(q='${fileName}')`;
    }
    
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}