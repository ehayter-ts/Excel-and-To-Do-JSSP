import '@k2oss/k2-broker-core';

metadata = {
    systemName: "MSExcelPlannerJSSP",
    displayName: "Microsoft Excel and Planner",
    description: "A connector for Microsoft Excel and Planner"
};

// Constants
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";

//
// Objects
const Drive = "drive";
const Excel = "excel";
const Group = "group";
const Planner = "planner";

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
const FolderName = "folderName";
const FolderPath = "folderPath";

const FileSearch = "searchFile";
const CreateFolder = "createFolder";

//
//Excel
const ExcelSheetName = "sheetName";
const Column1 = "column1";
const Column2 = "column2";
const Column3 = "column3";
const Column4 = "column4";
const Column5 = "column5";
const Column6 = "column6";
const Column7 = "column7";
const Column8 = "column8";
const Column9 = "column9";
const Column10 = "column10";
const Column11 = "column11";
const Column12 = "column12";
const Column13 = "column13";
const Column14 = "column14";
const Column15 = "column15";
const Column16 = "column16";
const Column17 = "column17";
const Column18 = "column18";
const Column19 = "column19";
const Column20 = "column20";

const UsedRangeItems = "getUsedRangeItems";

//
//Group
const GroupId = "groupId";
const GroupName = "groupName";
const GroupDescription = "groupDescription";
const GroupMail = "groupMail";
const GroupVisibility = "groupVisibility";
const GroupMailNickname = "groupMailNickname";
const GroupMailEnabled = "groupMailEnabled";
const GroupSecurityEnabled = "groupSecurityEnabled";

const GetGroups = "getGroups";
const CreateGroup = "createGroup";

//
//Planner
const PlanTitle = "planTitle";
const PlanOwnerGroup = "ownerGroup";
const PlanId = "planId";
const BucketName = "bucketName";
const BucketId = "bucketId";

const CreatePlan = "createPlan";
const CreateBucket = "createBucket";


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
                    },
                    [FolderName]: {
                        displayName: "Folder Name",
                        description: "Folder Name",
                        type: "string"
                    },
                    [FolderPath]: {
                        displayName: "Folder Path",
                        description: "Folder Path",
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
                    },
                    [CreateFolder]: {
                        displayName: "Create Folder",
                        type: "execute",
                        inputs: [FolderName, FolderPath],
                        requiredInputs: [FolderName],
                        outputs: []
                    }
                }
            },
            [Excel]: {
                displayName: "Excel",
                description: "Excel",
                properties: {
                    [FileId]: {
                        displayName: "File ID",
                        description: "File ID",
                        type: "string"
                    },
                    [ExcelSheetName]: {
                        displayName: "Sheet Name",
                        description: "Sheet Name",
                        type: "string"
                    },
                    [Column1]: {
                        displayName: "Column 1",
                        description: "Column 1",
                        type: "string"
                    },
                    [Column2]: {
                        displayName: "Column 2",
                        description: "Column 2",
                        type: "string"
                    },
                    [Column3]: {
                        displayName: "Column 3",
                        description: "Column 3",
                        type: "string"
                    },
                    [Column4]: {
                        displayName: "Column 4",
                        description: "Column 4",
                        type: "string"
                    },
                    [Column5]: {
                        displayName: "Column 5",
                        description: "Column 5",
                        type: "string"
                    },
                    [Column6]: {
                        displayName: "Column 6",
                        description: "Column 6",
                        type: "string"
                    },
                    [Column7]: {
                        displayName: "Column 7",
                        description: "Column 7",
                        type: "string"
                    },
                    [Column8]: {
                        displayName: "Column 8",
                        description: "Column 8",
                        type: "string"
                    },
                    [Column9]: {
                        displayName: "Column 9",
                        description: "Column 9",
                        type: "string"
                    },
                    [Column10]: {
                        displayName: "Column 10",
                        description: "Column 10",
                        type: "string"
                    },
                    [Column11]: {
                        displayName: "Column 11",
                        description: "Column 11",
                        type: "string"
                    },
                    [Column12]: {
                        displayName: "Column 12",
                        description: "Column 12",
                        type: "string"
                    },
                    [Column13]: {
                        displayName: "Column 13",
                        description: "Column 13",
                        type: "string"
                    },
                    [Column14]: {
                        displayName: "Column 14",
                        description: "Column 14",
                        type: "string"
                    },
                    [Column15]: {
                        displayName: "Column 15",
                        description: "Column 15",
                        type: "string"
                    },
                    [Column16]: {
                        displayName: "Column 16",
                        description: "Column 16",
                        type: "string"
                    },
                    [Column17]: {
                        displayName: "Column 17",
                        description: "Column 17",
                        type: "string"
                    },
                    [Column18]: {
                        displayName: "Column 18",
                        description: "Column 18",
                        type: "string"
                    },
                    [Column19]: {
                        displayName: "Column 19",
                        description: "Column 19",
                        type: "string"
                    },
                    [Column20]: {
                        displayName: "Column 20",
                        description: "Column 20",
                        type: "string"
                    }
                },
                methods: {
                    [UsedRangeItems]: {
                        displayName: "Get Worksheet Rows in Used Range",
                        type: "list",
                        inputs: [FileId, ExcelSheetName],
                        requiredInputs: [FileId, ExcelSheetName],
                        outputs: [Column1, Column2, Column3, Column4, Column5, Column6, Column7, Column8, Column9, Column10, Column11, Column12, Column13, Column14, Column15, Column16, Column17, Column18, Column19, Column20]
                    }
                }
            },
            [Group]: {
                displayName: "Group",
                description: "Group",
                properties: {
                    [GroupId]: {
                        displayName: "Group ID",
                        description: "Group ID",
                        type: "string"
                    },
                    [GroupName]: {
                        displayName: "Group Name",
                        description: "Group Name",
                        type: "string"
                    },
                    [GroupDescription]: {
                        displayName: "Group Description",
                        description: "Group Description",
                        type: "string"
                    },
                    [GroupMail]: {
                        displayName: "Group Mail",
                        description: "Group Mail",
                        type: "string"
                    },
                    [GroupVisibility]: {
                        displayName: "Group Visibility",
                        description: "Group Visibility",
                        type: "string"
                    },
                    [GroupMailEnabled]: {
                        displayName: "Mail Enabled",
                        description: "Mail Enabled",
                        type: "boolean"
                    },
                    [GroupMailNickname]: {
                        displayName: "Mail Nickname",
                        description: "Mail Nickname",
                        type: "string"
                    },
                    [GroupSecurityEnabled]: {
                        displayName: "Security Enabled",
                        description: "Security Enabled",
                        type: "boolean"
                    }
                },
                methods: {
                    [GetGroups]: {
                        displayName: "Get all Groups in Organisation",
                        type: "list",
                        inputs: [],
                        requiredInputs: [],
                        outputs: [GroupId, GroupName, GroupDescription, GroupMail, GroupVisibility]
                    },
                    [CreateGroup]: {
                        displayName: "Create Group",
                        type: "execute",
                        inputs: [GroupName, GroupDescription, GroupVisibility, GroupMailEnabled, GroupMailNickname, GroupSecurityEnabled],
                        requiredInputs: [GroupName, GroupMailEnabled, GroupMailNickname, GroupSecurityEnabled],
                        outputs: [GroupId, GroupName, GroupDescription, GroupMail, GroupVisibility]
                    }
                }
            },
            [Planner]: {
                displayName: "Planner",
                description: "Planner",
                properties: {
                    [PlanTitle]: {
                        displayName: "Planner Title",
                        description: "Planner Title",
                        type: "string"
                    },
                    [PlanOwnerGroup]: {
                        displayName: "Owner Group ID",
                        description: "Owner Group ID",
                        type: "string"
                    },
                    [PlanId]: {
                        displayName: "Plan ID",
                        description: "Plan ID",
                        type: "string"
                    },
                    [BucketName]: {
                        displayName: "Bucket Name",
                        description: "Bucket Name",
                        type: "string"
                    },
                    [BucketId]: {
                        displayName: "Bucket ID",
                        description: "Bucket ID",
                        type: "string"
                    }
                },
                methods: {
                    [CreatePlan]: {
                        displayName: "Create Plan",
                        type: "execute",
                        inputs: [PlanTitle, PlanOwnerGroup],
                        requiredInputs: [PlanTitle, PlanOwnerGroup],
                        outputs: [PlanId]
                    },
                    [CreateBucket]: {
                        displayName: "Create Bucket",
                        type: "execute",
                        inputs: [BucketName, PlanId],
                        requiredInputs: [BucketName, PlanId],
                        outputs: [BucketId]
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
        case Excel:
            onexecuteExcel(methodName, parameters, properties);
            break;
        case Group:
            onexecuteGroup(methodName, parameters, properties);
            break;
        case Planner:
            onexecutePlanner(methodName, parameters, properties);
            break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteDrive(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case FileSearch:
            onexecuteSearchFile(parameters, properties);
            break;
        case CreateFolder:
            onexecuteCreateFolder(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteExcel(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case UsedRangeItems:
            onexecuteUsedRange(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteGroup(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case GetGroups:
            onexecuteGetGroups(parameters, properties);
            break;
        case CreateGroup:
            onexecuteCreateGroup(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecutePlanner(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case CreatePlan:
            onexecuteCreatePlan(parameters, properties);
            break;
        case CreateBucket:
            onexecuteCreateBucket(parameters, properties);
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
            var obj;

            try {
                obj = JSON.parse(xhr.responseText);
            }
            catch (e) {
                //do nothing
            }

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

        if (a.value.length > 0) {
            resultObj[FileId] = a.value[0].id;
            resultObj[FileWebUrl] = a.value[0].webUrl;
            resultObj[FileSize] = a.value[0].size;
            resultObj[FileName] = a.value[0].name;
            resultObj[FileCreated] = Date.parse(a.value[0].createdDateTime);
            resultObj[FileCreatedBy] = a.value[0].createdBy.user.displayName;
            resultObj[FileCreatedByEmail] = a.value[0].createdBy.user.email;
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
    if ((typeof filePath === "string") && (filePath != "")) {
        url = baseUriEndpoint + `/me/drive/root:/${filePath}:/search(q='${fileName}')`;
    }
    else {
        url = baseUriEndpoint + `/me/drive/root/search(q='${fileName}')`;
    }

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteCreateFolder(parameters: SingleRecord, properties: SingleRecord) {
    CreateDriveFolder(parameters, properties, function () {
        //postResult({});
    });
}

function CreateDriveFolder(parameters: SingleRecord, properties: SingleRecord, cb) {
    let folderName = properties[FolderName];
    let folderPath = properties[FolderPath];

    if (!(typeof folderName === "string")) throw new Error("properties[FolderName] is not of type string");

    var url;
    if ((typeof folderPath === "string") && (folderPath != "")) {
        url = baseUriEndpoint + `/me/drive/root:/${folderPath}:/children'`;
    }
    else {
        url = baseUriEndpoint + `/me/drive/root/children`;
    }

    var data = {
        "name": folderName,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace"
    };

    ExecuteRequest(url, JSON.stringify(data), "POST", function () {
        if (typeof cb === 'function')
            cb();
    });
}

function onexecuteUsedRange(parameters: SingleRecord, properties: SingleRecord) {
    GetRangeItems(parameters, properties, function (a) {
        var obj = a.text.map(x => {
            var obj = {};

            for (var i = 0; i < x.length; i++) {
                if ((i + 1) < 21) {
                    obj["column" + (i + 1)] = x[i];
                }
            }
            return obj;
        });

        postResult(obj);
    });
}

function GetRangeItems(parameters: SingleRecord, properties: SingleRecord, cb) {
    let fileId = properties[FileId];
    let sheetName = properties[ExcelSheetName];

    if (!(typeof fileId === "string")) throw new Error("properties[FileId] is not of type string");
    if (!(typeof sheetName === "string")) throw new Error("properties[ExcelSheetName] is not of type string");

    var url = baseUriEndpoint + `/me/drive/items/${fileId}/workbook/worksheets('${sheetName}')/usedRange`;

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteGetGroups(parameters: SingleRecord, properties: SingleRecord) {
    GetGroupItems(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [GroupId]: x.id,
                [GroupName]: x.displayName,
                [GroupDescription]: x.description,
                [GroupMail]: x.mail,
                [GroupVisibility]: x.visibility
            };
        }));
    });
}

function GetGroupItems(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + '/groups';

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteCreatePlan(parameters: SingleRecord, properties: SingleRecord) {
    CreatePlannerPlan(parameters, properties, function (a) {
        postResult({
            [PlanId]: a.id
        });
    });
}

function CreatePlannerPlan(parameters: SingleRecord, properties: SingleRecord, cb) {
    let planTitle = properties[PlanTitle];
    let planGroup = properties[PlanOwnerGroup];

    if (!(typeof planTitle === "string")) throw new Error("properties[PlanTitle] is not of type string");
    if (!(typeof planGroup === "string")) throw new Error("properties[PlanOwnerGroup] is not of type string");

    var url = baseUriEndpoint + '/planner/plans';

    var data = {
        title: planTitle,
        owner: planGroup
    };

    ExecuteRequest(url, JSON.stringify(data), "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteCreateBucket(parameters: SingleRecord, properties: SingleRecord) {
    CreatePlannerPlanBucket(parameters, properties, function (a) {
        postResult({
            [BucketId]: a.id
        });
    });
}

function CreatePlannerPlanBucket(parameters: SingleRecord, properties: SingleRecord, cb) {
    let bucketName = properties[BucketName];
    let planId = properties[PlanId];

    if (!(typeof bucketName === "string")) throw new Error("properties[BucketName] is not of type string");
    if (!(typeof planId === "string")) throw new Error("properties[PlanId] is not of type string");

    var url = baseUriEndpoint + '/planner/buckets';

    var data = {
        name: bucketName,
        planId: planId
    };

    ExecuteRequest(url, JSON.stringify(data), "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteCreateGroup(parameters: SingleRecord, properties: SingleRecord) {
    CreateNewGroup(parameters, properties, function (a) {
        postResult({
            [GroupId]: a.id,
            [GroupName]: a.displayName,
            [GroupDescription]: a.description,
            [GroupMail]: a.mail,
            [GroupVisibility]: a.visibility
        });
    });
}

function CreateNewGroup(parameters: SingleRecord, properties: SingleRecord, cb) {
    let groupName = properties[GroupName];
    let groupDesc = properties[GroupDescription];
    let groupVisibility = properties[GroupVisibility];
    let groupMailEnabled = properties[GroupMailEnabled];
    let groupMailNic = properties[GroupMailNickname];
    let groupSecurityEnabled = properties[GroupSecurityEnabled];


    if (!(typeof groupName === "string")) throw new Error("properties[GroupName] is not of type string");
    if (!(typeof groupMailNic === "string")) throw new Error("properties[GroupMailNickname] is not of type string");

    var url = baseUriEndpoint + '/groups';

    var data = {
        displayName: groupName,
        mailEnabled: groupMailEnabled,
        mailNickname: groupMailNic,
        securityEnabled: groupSecurityEnabled,
        description: groupDesc,
        visibility: (groupVisibility != null && groupVisibility != "") ? groupVisibility : "Public"
    };

    ExecuteRequest(url, JSON.stringify(data), "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}