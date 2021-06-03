!function(){metadata={systemName:"MSExcelPlannerJSSP",displayName:"Microsoft Excel and Planner",description:"A connector for Microsoft Excel and Planner"};const e="https://graph.microsoft.com/v1.0",t="fileId",i="planId";function r(e,t,i,r){var o=new XMLHttpRequest;o.onreadystatechange=function(){if(console.log("ExecuteRequest XHR status: "+o.status+","+o.responseText),4===o.readyState)if(201==o.status){try{e=JSON.parse(o.responseText)}catch(e){}"function"==typeof r&&r(e)}else if(204==o.status)"function"==typeof r&&r(o.responseText);else if(200==o.status){var e=JSON.parse(o.responseText);"function"==typeof r&&r(e)}else if(202==o.status)"function"==typeof r&&r(null);else{if(400==o.status){e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if(404!=o.status){postResult({});e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if("No Team found with Group id"!=(e=JSON.parse(o.responseText)).error.message.startswith)throw new Error(e.error.code+": "+e.error.message)}},console.log("MSTeamsConnector ExecuteRequest: "+e),o.open(i.toUpperCase(),e),o.withCredentials=!0,o.setRequestHeader("Accept","application/json"),"PUT"!=i.toUpperCase()&&"POST"!=i.toUpperCase()&&"PATCH"!=i.toUpperCase()||o.setRequestHeader("Content-Type","application/json"),o.send(t)}function o(e,i,o){let s=i[t],n=i.sheetName,a=i.worksheetRange;if("string"!=typeof s)throw new Error("properties[FileId] is not of type string");if("string"!=typeof n)throw new Error("properties[ExcelSheetName] is not of type string");if("string"!=typeof a)throw new Error("properties[WorksheetRange] is not of type string");r(`https://graph.microsoft.com/v1.0/me/drive/items/${s}/workbook/worksheets/${n}/range(address='${a}')`,null,"GET",(function(e){"function"==typeof o&&o(e)}))}function s(e,i,o){let s=i[t],n=i.sheetName;if("string"!=typeof s)throw new Error("properties[FileId] is not of type string");if("string"!=typeof n)throw new Error("properties[ExcelSheetName] is not of type string");r(`https://graph.microsoft.com/v1.0/me/drive/items/${s}/workbook/worksheets('${n}')/usedRange`,null,"GET",(function(e){"function"==typeof o&&o(e)}))}ondescribe=function(){postSchema({objects:{drive:{displayName:"Drive",description:"Drive",properties:{[t]:{displayName:"File ID",description:"File ID",type:"string"},fileWebUrl:{displayName:"File Web URL",description:"File Web URL",type:"string"},fileSize:{displayName:"File Size",description:"File Size",type:"number"},fileName:{displayName:"File Name",description:"File Name",type:"string"},fileCreatedDate:{displayName:"File Created Date",description:"File Created Date",type:"dateTime"},fileCreatedBy:{displayName:"File Created By",description:"File Created By",type:"string"},fileCreatedByEmail:{displayName:"File Created By Email",description:"File Created By Email",type:"string"},fileMimeType:{displayName:"File Mime Type",description:"File Mime Type",type:"string"},filePath:{displayName:"File Path",description:"File Path",type:"string"},folderName:{displayName:"Folder Name",description:"Folder Name",type:"string"},folderPath:{displayName:"Folder Path",description:"Folder Path",type:"string"},file:{displayName:"File",description:"File",type:"attachment"}},methods:{searchFile:{displayName:"Search for a file in my OneDrive",type:"list",inputs:["fileName","filePath"],requiredInputs:["fileName"],outputs:[t,"fileWebUrl","fileSize","fileName","fileCreatedDate","fileCreatedBy","fileCreatedByEmail","fileMimeType"]},createFolder:{displayName:"Create Folder",type:"execute",inputs:["folderName","folderPath"],requiredInputs:["folderName"],outputs:[]}}},excel:{displayName:"Excel",description:"Excel",properties:{[t]:{displayName:"File ID",description:"File ID",type:"string"},sheetName:{displayName:"Sheet Name",description:"Sheet Name",type:"string"},column1:{displayName:"Column 1",description:"Column 1",type:"string"},column2:{displayName:"Column 2",description:"Column 2",type:"string"},column3:{displayName:"Column 3",description:"Column 3",type:"string"},column4:{displayName:"Column 4",description:"Column 4",type:"string"},column5:{displayName:"Column 5",description:"Column 5",type:"string"},column6:{displayName:"Column 6",description:"Column 6",type:"string"},column7:{displayName:"Column 7",description:"Column 7",type:"string"},column8:{displayName:"Column 8",description:"Column 8",type:"string"},column9:{displayName:"Column 9",description:"Column 9",type:"string"},column10:{displayName:"Column 10",description:"Column 10",type:"string"},column11:{displayName:"Column 11",description:"Column 11",type:"string"},column12:{displayName:"Column 12",description:"Column 12",type:"string"},column13:{displayName:"Column 13",description:"Column 13",type:"string"},column14:{displayName:"Column 14",description:"Column 14",type:"string"},column15:{displayName:"Column 15",description:"Column 15",type:"string"},column16:{displayName:"Column 16",description:"Column 16",type:"string"},column17:{displayName:"Column 17",description:"Column 17",type:"string"},column18:{displayName:"Column 18",description:"Column 18",type:"string"},column19:{displayName:"Column 19",description:"Column 19",type:"string"},column20:{displayName:"Column 20",description:"Column 20",type:"string"},columnName:{displayName:"Column Count",description:"Column Count",type:"string"},worksheetId:{displayName:"Worksheet ID",description:"Worksheet ID",type:"string"},worksheetName:{displayName:"Worksheet Name",description:"Worksheet Name",type:"string"},worksheetPosition:{displayName:"Worksheet Position",description:"Worksheet Position",type:"string"},worksheetVisibility:{displayName:"Worksheet Visibility",description:"Worksheet Visibility",type:"string"},worksheetRange:{displayName:"Worksheet Range",description:"Worksheet Range",type:"string"},rowCount:{displayName:"Row Count",description:"Row Count",type:"number"},columnCount:{displayName:"Column Count",description:"Column Count",type:"number"},cellValue:{displayName:"Cell Value",description:"Cell Value",type:"string"}},methods:{getUsedRange:{displayName:"Get Used Range",type:"read",inputs:[t,"sheetName"],requiredInputs:[t,"sheetName"],outputs:["worksheetRange"]},getUsedRangeItems:{displayName:"Get Worksheet Rows in Used Range",type:"list",inputs:[t,"sheetName"],requiredInputs:[t,"sheetName"],outputs:["column1","column2","column3","column4","column5","column6","column7","column8","column9","column10","column11","column12","column13","column14","column15","column16","column17","column18","column19","column20"]},getUsedRangeColumnNames:{displayName:"Get Used Range Column Counts",type:"list",inputs:[t,"sheetName"],requiredInputs:[t,"sheetName"],outputs:["columnName"]},getRangeColumnNames:{displayName:"Get Range Column Counts",type:"list",inputs:[t,"sheetName","worksheetRange"],requiredInputs:[t,"sheetName","worksheetRange"],outputs:["columnName"]},getWorkbookWorksheets:{displayName:"Get Worksheets",type:"list",inputs:[t],requiredInputs:[t],outputs:["worksheetId","worksheetName","worksheetPosition","worksheetVisibility"]},getRangeItems:{displayName:"Get Worksheet Rows in a Specified Range",type:"list",inputs:[t,"sheetName","worksheetRange"],requiredInputs:[t,"sheetName","worksheetRange"],outputs:["column1","column2","column3","column4","column5","column6","column7","column8","column9","column10","column11","column12","column13","column14","column15","column16","column17","column18","column19","column20"]},getCell:{displayName:"Get a Cell Value",type:"read",inputs:[t,"sheetName","worksheetRange","rowCount","columnCount"],requiredInputs:[t,"sheetName","worksheetRange","rowCount","columnCount"],outputs:["cellValue"]}}},group:{displayName:"Group",description:"Group",properties:{groupId:{displayName:"Group ID",description:"Group ID",type:"string"},groupName:{displayName:"Group Name",description:"Group Name",type:"string"},groupDescription:{displayName:"Group Description",description:"Group Description",type:"string"},groupMail:{displayName:"Group Mail",description:"Group Mail",type:"string"},groupVisibility:{displayName:"Group Visibility",description:"Group Visibility",type:"string"},groupMailEnabled:{displayName:"Mail Enabled",description:"Mail Enabled",type:"boolean"},groupMailNickname:{displayName:"Mail Nickname",description:"Mail Nickname",type:"string"},groupSecurityEnabled:{displayName:"Security Enabled",description:"Security Enabled",type:"boolean"},groupOwnerId:{displayName:"Group Owner ID",description:"Group Owner ID",type:"string"},userId:{displayName:"User ID",description:"User ID",type:"string"}},methods:{getGroups:{displayName:"Get all Groups in Organisation",type:"list",inputs:[],requiredInputs:[],outputs:["groupId","groupName","groupDescription","groupMail","groupVisibility"]},createGroup:{displayName:"Create Group",type:"execute",inputs:["groupName","groupDescription","groupVisibility","groupMailEnabled","groupMailNickname","groupSecurityEnabled","groupOwnerId"],requiredInputs:["groupName","groupMailEnabled","groupMailNickname","groupSecurityEnabled","groupOwnerId"],outputs:["groupId","groupName","groupDescription","groupMail","groupVisibility"]},addMemberToGroup:{displayName:"Add Member to Group",type:"execute",inputs:["groupId","userId"],requiredInputs:["groupId","userId"],outputs:[]}}},planner:{displayName:"Planner",description:"Planner",properties:{planTitle:{displayName:"Planner Title",description:"Planner Title",type:"string"},ownerGroup:{displayName:"Owner Group ID",description:"Owner Group ID",type:"string"},[i]:{displayName:"Plan ID",description:"Plan ID",type:"string"},bucketName:{displayName:"Bucket Name",description:"Bucket Name",type:"string"},bucketId:{displayName:"Bucket ID",description:"Bucket ID",type:"string"},taskTitle:{displayName:"Task Title",description:"Task Title",type:"string"},taskUserId:{displayName:"Task Assigned To User ID",description:"Task Assigned To User ID",type:"string"},taskDueDate:{displayName:"Task Due Date",description:"Task Due Date",type:"dateTime"},taskId:{displayName:"Task ID",description:"Task ID",type:"string"}},methods:{createPlan:{displayName:"Create Plan",type:"execute",inputs:["planTitle","ownerGroup"],requiredInputs:["planTitle","ownerGroup"],outputs:[i]},createBucket:{displayName:"Create Bucket",type:"execute",inputs:["bucketName",i],requiredInputs:["bucketName",i],outputs:["bucketId"]},getGroupPlans:{displayName:"Get Group Plans",type:"list",inputs:["ownerGroup"],requiredInputs:["ownerGroup"],outputs:[i,"planTitle"]},getPlanBuckets:{displayName:"Get Plan Buckets",type:"list",inputs:[i],requiredInputs:[i],outputs:["bucketId","bucketName"]},createTask:{displayName:"Create Task",type:"execute",inputs:[i,"taskTitle","taskUserId","bucketId","taskDueDate"],requiredInputs:[i,"taskTitle","taskUserId"],outputs:["taskId"]}}},user:{displayName:"User",description:"User",properties:{userEmail:{displayName:"User Email",description:"User Email",type:"string"},userId:{displayName:"User ID",description:"User ID",type:"string"},userDisplayName:{displayName:"User Display Name",description:"User Display Name",type:"string"},userGivenName:{displayName:"User Given Name",description:"User Given Name",type:"string"},userLastName:{displayName:"User Last Name",description:"User Last Name",type:"string"},userJobTitle:{displayName:"User Job Title",description:"User Job Title",type:"string"}},methods:{getUserByEmail:{displayName:"Get User by Email",type:"read",inputs:["userEmail"],requiredInputs:["userEmail"],outputs:["userId","userDisplayName","userEmail","userGivenName","userLastName","userJobTitle"]}}}}})},onexecute=function({objectName:n,methodName:a,parameters:p,properties:l}){switch(n){case"drive":!function(i,o,s){switch(i){case"searchFile":!function(e,i){!function(e,t,i){let o=t.fileName,s=t.filePath;if("string"!=typeof o)throw new Error("properties[FileName] is not of type string");var n;n="string"==typeof s&&""!=s?`https://graph.microsoft.com/v1.0/me/drive/root:/${s}:/search(q='${o}')`:`https://graph.microsoft.com/v1.0/me/drive/root/search(q='${o}')`;r(n,null,"GET",(function(e){"function"==typeof i&&i(e)}))}(0,i,(function(e){var i={[t]:"",fileWebUrl:"",fileSize:0,fileName:"",fileCreatedDate:null,fileCreatedBy:"",fileCreatedByEmail:"",fileMimeType:""};e.value.length>0&&(i[t]=e.value[0].id,i.fileWebUrl=e.value[0].webUrl,i.fileSize=e.value[0].size,i.fileName=e.value[0].name,i.fileCreatedDate=Date.parse(e.value[0].createdDateTime),i.fileCreatedBy=e.value[0].createdBy.user.displayName,i.fileCreatedByEmail=e.value[0].createdBy.user.email,i.fileMimeType=e.value[0].file.mimeType),postResult(i)}))}(0,s);break;case"createFolder":!function(t,i){!function(t,i,o){let s=i.folderName,n=i.folderPath;if("string"!=typeof s)throw new Error("properties[FolderName] is not of type string");var a;a="string"==typeof n&&""!=n?`https://graph.microsoft.com/v1.0/me/drive/root:/${n}:/children'`:e+"/me/drive/root/children";var p={name:s,folder:{},"@microsoft.graph.conflictBehavior":"replace"};r(a,JSON.stringify(p),"POST",(function(){"function"==typeof o&&o()}))}(0,i,(function(){}))}(0,s);break;default:throw new Error("The method "+i+" is not supported..")}}(a,0,l);break;case"excel":!function(e,i,n){switch(e){case"getUsedRange":!function(e,t){s(e,t,(function(e){postResult({worksheetRange:e.address.split("!")[1]})}))}(i,n);break;case"getUsedRangeItems":!function(e,t){s(e,t,(function(e){var t=e.values.map(e=>{for(var t={},i=0;i<e.length;i++)i+1<21&&(t["column"+(i+1)]=e[i]);return t});postResult(t)}))}(i,n);break;case"getUsedRangeColumnNames":!function(e,t){s(e,t,(function(e){var t=0,i=[];if(e.text.length>0){t=e.text[0].length;for(var r=1;r<t+1;r++)i.push({columnName:r})}postResult(i)}))}(i,n);break;case"getRangeColumnNames":!function(e,t){o(e,t,(function(e){var t=0,i=[];if(e.text.length>0){t=e.text[0].length;for(var r=1;r<t+1;r++)i.push({columnName:r})}postResult(i)}))}(i,n);break;case"getWorkbookWorksheets":!function(e,i){!function(e,i,o){let s=i[t];if("string"!=typeof s)throw new Error("properties[FileId] is not of type string");r(`https://graph.microsoft.com/v1.0/me/drive/items/${s}/workbook/worksheets`,null,"GET",(function(e){"function"==typeof o&&o(e)}))}(0,i,(function(e){var t=e.value.map(e=>({worksheetId:e.id,worksheetName:e.name,worksheetPosition:e.position,worksheetVisibility:e.visibility}));postResult(t)}))}(0,n);break;case"getRangeItems":!function(e,t){o(e,t,(function(e){var t=e.values.map(e=>{for(var t={},i=0;i<e.length;i++)i+1<21&&(t["column"+(i+1)]=e[i]);return t});postResult(t)}))}(i,n);break;case"getCell":!function(e,i){!function(e,i,o){let s=i[t],n=i.sheetName,a=i.worksheetRange,p=i.rowCount,l=i.columnCount;if("string"!=typeof s)throw new Error("properties[FileId] is not of type string");if("string"!=typeof n)throw new Error("properties[ExcelSheetName] is not of type string");if("string"!=typeof a)throw new Error("properties[WorksheetRange] is not of type string");if("number"!=typeof p)throw new Error("properties[RowCount] is not of type integer");if("number"!=typeof l)throw new Error("properties[ColumnCount] is not of type integer");r(`https://graph.microsoft.com/v1.0/me/drive/items/${s}/workbook/worksheets/${n}/range(address='${a}')/cell(row=${p},column=${l})`,null,"GET",(function(e){"function"==typeof o&&o(e)}))}(0,i,(function(e){postResult({cellValue:e.values[0][0]})}))}(0,n);break;default:throw new Error("The method "+e+" is not supported..")}}(a,p,l);break;case"group":!function(t,i,o){switch(t){case"getGroups":s=function(e){postResult(e.value.map(e=>({groupId:e.id,groupName:e.displayName,groupDescription:e.description,groupMail:e.mail,groupVisibility:e.visibility})))},r(e+"/groups",null,"GET",(function(e){"function"==typeof s&&s(e)}));break;case"createGroup":!function(t,i){!function(t,i,o){let s=i.groupName,n=i.groupDescription,a=i.groupVisibility,p=i.groupMailEnabled,l=i.groupMailNickname,u=i.groupSecurityEnabled,m=i.groupOwnerId;if("string"!=typeof s)throw new Error("properties[GroupName] is not of type string");if("string"!=typeof l)throw new Error("properties[GroupMailNickname] is not of type string");if("string"!=typeof m)throw new Error("properties[GroupOwnerId] is not of type string");var c={displayName:s,mailEnabled:p,mailNickname:l,securityEnabled:u,description:n,visibility:null!=a&&""!=a?a:"Public",groupTypes:["Unified"],"owners@odata.bind":["https://graph.microsoft.com/v1.0/users/"+m]};r(e+"/groups",JSON.stringify(c),"POST",(function(e){"function"==typeof o&&o(e)}))}(0,i,(function(e){postResult({groupId:e.id,groupName:e.displayName,groupDescription:e.description,groupMail:e.mail,groupVisibility:e.visibility})}))}(0,o);break;case"addMemberToGroup":!function(e,t){!function(e,t,i){let o=t.groupId,s=t.userId;if("string"!=typeof o)throw new Error("properties[GroupId] is not of type string");if("string"!=typeof s)throw new Error("properties[UserId] is not of type string");var n={"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/"+s};r(`https://graph.microsoft.com/v1.0/groups/${o}/members/$ref`,JSON.stringify(n),"POST",(function(e){"function"==typeof i&&i(e)}))}(0,t,(function(e){postResult({})}))}(0,o);break;default:throw new Error("The method "+t+" is not supported..")}var s}(a,0,l);break;case"planner":!function(t,o,s){switch(t){case"createPlan":!function(t,o){!function(t,i,o){let s=i.planTitle,n=i.ownerGroup;if("string"!=typeof s)throw new Error("properties[PlanTitle] is not of type string");if("string"!=typeof n)throw new Error("properties[PlanOwnerGroup] is not of type string");var a={title:s,owner:n};r(e+"/planner/plans",JSON.stringify(a),"POST",(function(e){"function"==typeof o&&o(e)}))}(0,o,(function(e){postResult({[i]:e.id})}))}(0,s);break;case"createBucket":!function(t,o){!function(t,o,s){let n=o.bucketName,a=o[i];if("string"!=typeof n)throw new Error("properties[BucketName] is not of type string");if("string"!=typeof a)throw new Error("properties[PlanId] is not of type string");var p={name:n,planId:a};r(e+"/planner/buckets",JSON.stringify(p),"POST",(function(e){"function"==typeof s&&s(e)}))}(0,o,(function(e){postResult({bucketId:e.id})}))}(0,s);break;case"getGroupPlans":!function(e,t){!function(e,t,i){let o=t.ownerGroup;if("string"!=typeof o)throw new Error("properties[PlanOwnerGroup] is not of type string");r(`https://graph.microsoft.com/v1.0/groups/${o}/planner/plans`,null,"GET",(function(e){"function"==typeof i&&i(e)}))}(0,t,(function(e){postResult(e.value.map(e=>({[i]:e.id,planTitle:e.title})))}))}(0,s);break;case"getPlanBuckets":!function(e,t){!function(e,t,o){let s=t[i];if("string"!=typeof s)throw new Error("properties[PlanId] is not of type string");r(`https://graph.microsoft.com/v1.0/planner/plans/${s}/buckets`,null,"GET",(function(e){"function"==typeof o&&o(e)}))}(0,t,(function(e){postResult(e.value.map(e=>({bucketId:e.id,bucketName:e.name})))}))}(0,s);break;case"createTask":!function(t,o){!function(t,o,s){let n=o[i],a=o.taskUserId,p=o.taskTitle,l=null!=o.taskDueDate&&""!=o.taskDueDate?o.taskDueDate.toString():"",u=o.bucketId;""!=l&&(l=new Date(l).toISOString().split("T")[0]);if("string"!=typeof n)throw new Error("properties[PlanId] is not of type string");if("string"!=typeof a)throw new Error("properties[TaskUserId] is not of type string");if("string"!=typeof p)throw new Error("properties[TaskTitle] is not of type string");r(e+"/planner/tasks",`{"planId":"${n}","title":"${p}","dueDateTime":"${l}","bucketId":"${u}","assignments":{"${a}":{"@odata.type":"#microsoft.graph.plannerAssignment","orderHint":" !"}}}`,"POST",(function(e){"function"==typeof s&&s(e)}))}(0,o,(function(e){postResult({taskId:e.id})}))}(0,s);break;default:throw new Error("The method "+t+" is not supported..")}}(a,0,l);break;case"user":!function(t,i,o){switch(t){case"getUserByEmail":!function(t,i){!function(t,i,o){let s=i.userEmail;if("string"!=typeof s)throw new Error("properties[UserEmail] is not of type string");r(e+"/users/"+s,null,"GET",(function(e){"function"==typeof o&&o(e)}))}(0,i,(function(e){postResult({userId:e.id,userDisplayName:e.displayName,userEmail:e.mail,userGivenName:e.givenName,userLastName:e.surname,userJobTitle:e.jobTitle})}))}(0,o);break;default:throw new Error("The method "+t+" is not supported..")}}(a,0,l);break;default:throw new Error("The object "+n+" is not supported.")}}}();
//# sourceMappingURL=index.js.map
