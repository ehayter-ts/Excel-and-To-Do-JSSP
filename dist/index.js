!function(){metadata={systemName:"MSExcelPlannerJSSP",displayName:"Microsoft Excel and Planner",description:"A connector for Microsoft Excel and Planner"};const e="https://graph.microsoft.com/v1.0";function i(e,i,t,r){var o=new XMLHttpRequest;o.onreadystatechange=function(){if(console.log("ExecuteRequest XHR status: "+o.status+","+o.responseText),4===o.readyState)if(201==o.status){try{e=JSON.parse(o.responseText)}catch(e){}"function"==typeof r&&r(e)}else if(204==o.status)"function"==typeof r&&r(o.responseText);else if(200==o.status){var e=JSON.parse(o.responseText);"function"==typeof r&&r(e)}else if(202==o.status)"function"==typeof r&&r(null);else{if(400==o.status){e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if(404!=o.status){postResult({});e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if("No Team found with Group id"!=(e=JSON.parse(o.responseText)).error.message.startswith)throw new Error(e.error.code+": "+e.error.message)}},console.log("MSTeamsConnector ExecuteRequest: "+e),o.open(t.toUpperCase(),e),o.withCredentials=!0,o.setRequestHeader("Accept","application/json"),"PUT"!=t.toUpperCase()&&"POST"!=t.toUpperCase()&&"PATCH"!=t.toUpperCase()||o.setRequestHeader("Content-Type","application/json"),o.send(i)}ondescribe=function(){postSchema({objects:{drive:{displayName:"Drive",description:"Drive",properties:{fileId:{displayName:"File ID",description:"File ID",type:"string"},fileWebUrl:{displayName:"File Web URL",description:"File Web URL",type:"string"},fileSize:{displayName:"File Size",description:"File Size",type:"number"},fileName:{displayName:"File Name",description:"File Name",type:"string"},fileCreatedDate:{displayName:"File Created Date",description:"File Created Date",type:"dateTime"},fileCreatedBy:{displayName:"File Created By",description:"File Created By",type:"string"},fileCreatedByEmail:{displayName:"File Created By Email",description:"File Created By Email",type:"string"},fileMimeType:{displayName:"File Mime Type",description:"File Mime Type",type:"string"},filePath:{displayName:"File Path",description:"File Path",type:"string"},folderName:{displayName:"Folder Name",description:"Folder Name",type:"string"},folderPath:{displayName:"Folder Path",description:"Folder Path",type:"string"}},methods:{searchFile:{displayName:"Search for a file in my OneDrive",type:"read",inputs:["fileName","filePath"],requiredInputs:["fileName"],outputs:["fileId","fileWebUrl","fileSize","fileName","fileCreatedDate","fileCreatedBy","fileCreatedByEmail","fileMimeType"]},createFolder:{displayName:"Create Folder",type:"execute",inputs:["folderName","folderPath"],requiredInputs:["folderName"],outputs:[]}}},excel:{displayName:"Excel",description:"Excel",properties:{fileId:{displayName:"File ID",description:"File ID",type:"string"},sheetName:{displayName:"Sheet Name",description:"Sheet Name",type:"string"},column1:{displayName:"Column 1",description:"Column 1",type:"string"},column2:{displayName:"Column 2",description:"Column 2",type:"string"},column3:{displayName:"Column 3",description:"Column 3",type:"string"},column4:{displayName:"Column 4",description:"Column 4",type:"string"},column5:{displayName:"Column 5",description:"Column 5",type:"string"},column6:{displayName:"Column 6",description:"Column 6",type:"string"},column7:{displayName:"Column 7",description:"Column 7",type:"string"},column8:{displayName:"Column 8",description:"Column 8",type:"string"},column9:{displayName:"Column 9",description:"Column 9",type:"string"},column10:{displayName:"Column 10",description:"Column 10",type:"string"},column11:{displayName:"Column 11",description:"Column 11",type:"string"},column12:{displayName:"Column 12",description:"Column 12",type:"string"},column13:{displayName:"Column 13",description:"Column 13",type:"string"},column14:{displayName:"Column 14",description:"Column 14",type:"string"},column15:{displayName:"Column 15",description:"Column 15",type:"string"},column16:{displayName:"Column 16",description:"Column 16",type:"string"},column17:{displayName:"Column 17",description:"Column 17",type:"string"},column18:{displayName:"Column 18",description:"Column 18",type:"string"},column19:{displayName:"Column 19",description:"Column 19",type:"string"},column20:{displayName:"Column 20",description:"Column 20",type:"string"}},methods:{getUsedRangeItems:{displayName:"Get Worksheet Rows in Used Range",type:"list",inputs:["fileId","sheetName"],requiredInputs:["fileId","sheetName"],outputs:["column1","column2","column3","column4","column5","column6","column7","column8","column9","column10","column11","column12","column13","column14","column15","column16","column17","column18","column19","column20"]}}},group:{displayName:"Group",description:"Group",properties:{groupId:{displayName:"Group ID",description:"Group ID",type:"string"},groupName:{displayName:"Group Name",description:"Group Name",type:"string"},groupDescription:{displayName:"Group Description",description:"Group Description",type:"string"},groupMail:{displayName:"Group Mail",description:"Group Mail",type:"string"},groupVisibility:{displayName:"Group Visibility",description:"Group Visibility",type:"string"},groupMailEnabled:{displayName:"Mail Enabled",description:"Mail Enabled",type:"boolean"},groupMailNickname:{displayName:"Mail Nickname",description:"Mail Nickname",type:"string"},groupSecurityEnabled:{displayName:"Security Enabled",description:"Security Enabled",type:"boolean"}},methods:{getGroups:{displayName:"Get all Groups in Organisation",type:"list",inputs:[],requiredInputs:[],outputs:["groupId","groupName","groupDescription","groupMail","groupVisibility"]},createGroup:{displayName:"Create Group",type:"execute",inputs:["groupName","groupDescription","groupVisibility","groupMailEnabled","groupMailNickname","groupSecurityEnabled"],requiredInputs:["groupName","groupMailEnabled","groupMailNickname","groupSecurityEnabled"],outputs:["groupId","groupName","groupDescription","groupMail","groupVisibility"]}}},planner:{displayName:"Planner",description:"Planner",properties:{planTitle:{displayName:"Planner Title",description:"Planner Title",type:"string"},ownerGroup:{displayName:"Owner Group ID",description:"Owner Group ID",type:"string"},planId:{displayName:"Plan ID",description:"Plan ID",type:"string"},bucketName:{displayName:"Bucket Name",description:"Bucket Name",type:"string"},bucketId:{displayName:"Bucket ID",description:"Bucket ID",type:"string"}},methods:{createPlan:{displayName:"Create Plan",type:"execute",inputs:["planTitle","ownerGroup"],requiredInputs:["planTitle","ownerGroup"],outputs:["planId"]},createBucket:{displayName:"Create Bucket",type:"execute",inputs:["bucketName","planId"],requiredInputs:["bucketName","planId"],outputs:["bucketId"]}}}}})},onexecute=function({objectName:t,methodName:r,parameters:o,properties:n}){switch(t){case"drive":!function(t,r,o){switch(t){case"searchFile":!function(e,t){!function(e,t,r){let o=t.fileName,n=t.filePath;if("string"!=typeof o)throw new Error("properties[FileName] is not of type string");var a;a="string"==typeof n&&""!=n?`https://graph.microsoft.com/v1.0/me/drive/root:/${n}:/search(q='${o}')`:`https://graph.microsoft.com/v1.0/me/drive/root/search(q='${o}')`;i(a,null,"GET",(function(e){"function"==typeof r&&r(e)}))}(0,t,(function(e){var i={fileId:"",fileWebUrl:"",fileSize:0,fileName:"",fileCreatedDate:null,fileCreatedBy:"",fileCreatedByEmail:"",fileMimeType:""};e.value.length>0&&(i.fileId=e.value[0].id,i.fileWebUrl=e.value[0].webUrl,i.fileSize=e.value[0].size,i.fileName=e.value[0].name,i.fileCreatedDate=Date.parse(e.value[0].createdDateTime),i.fileCreatedBy=e.value[0].createdBy.user.displayName,i.fileCreatedByEmail=e.value[0].createdBy.user.email,i.fileMimeType=e.value[0].file.mimeType),postResult(i)}))}(0,o);break;case"createFolder":!function(t,r){!function(t,r,o){let n=r.folderName,a=r.folderPath;if("string"!=typeof n)throw new Error("properties[FolderName] is not of type string");var p;p="string"==typeof a&&""!=a?`https://graph.microsoft.com/v1.0/me/drive/root:/${a}:/children'`:e+"/me/drive/root/children";var s={name:n,folder:{},"@microsoft.graph.conflictBehavior":"replace"};i(p,JSON.stringify(s),"POST",(function(){"function"==typeof o&&o()}))}(0,r,(function(){}))}(0,o);break;default:throw new Error("The method "+t+" is not supported..")}}(r,0,n);break;case"excel":!function(e,t,r){switch(e){case"getUsedRangeItems":!function(e,t){!function(e,t,r){let o=t.fileId,n=t.sheetName;if("string"!=typeof o)throw new Error("properties[FileId] is not of type string");if("string"!=typeof n)throw new Error("properties[ExcelSheetName] is not of type string");i(`https://graph.microsoft.com/v1.0/me/drive/items/${o}/workbook/worksheets('${n}')/usedRange`,null,"GET",(function(e){"function"==typeof r&&r(e)}))}(0,t,(function(e){var i=e.text.map(e=>{for(var i={},t=0;t<e.length;t++)t+1<21&&(i["column"+(t+1)]=e[t]);return i});postResult(i)}))}(0,r);break;default:throw new Error("The method "+e+" is not supported..")}}(r,0,n);break;case"group":!function(t,r,o){switch(t){case"getGroups":n=function(e){postResult(e.value.map(e=>({groupId:e.id,groupName:e.displayName,groupDescription:e.description,groupMail:e.mail,groupVisibility:e.visibility})))},i(e+"/groups",null,"GET",(function(e){"function"==typeof n&&n(e)}));break;case"createGroup":!function(t,r){!function(t,r,o){let n=r.groupName,a=r.groupDescription,p=r.groupVisibility,s=r.groupMailEnabled,l=r.groupMailNickname,u=r.groupSecurityEnabled;if("string"!=typeof n)throw new Error("properties[GroupName] is not of type string");if("string"!=typeof l)throw new Error("properties[GroupMailNickname] is not of type string");var c={displayName:n,mailEnabled:s,mailNickname:l,securityEnabled:u,description:a,visibility:null!=p&&""!=p?p:"Public"};i(e+"/groups",JSON.stringify(c),"POST",(function(e){"function"==typeof o&&o(e)}))}(0,r,(function(e){postResult({groupId:e.id,groupName:e.displayName,groupDescription:e.description,groupMail:e.mail,groupVisibility:e.visibility})}))}(0,o);break;default:throw new Error("The method "+t+" is not supported..")}var n}(r,0,n);break;case"planner":!function(t,r,o){switch(t){case"createPlan":!function(t,r){!function(t,r,o){let n=r.planTitle,a=r.ownerGroup;if("string"!=typeof n)throw new Error("properties[PlanTitle] is not of type string");if("string"!=typeof a)throw new Error("properties[PlanOwnerGroup] is not of type string");var p={title:n,owner:a};i(e+"/planner/plans",JSON.stringify(p),"POST",(function(e){"function"==typeof o&&o(e)}))}(0,r,(function(e){postResult({planId:e.id})}))}(0,o);break;case"createBucket":!function(t,r){!function(t,r,o){let n=r.bucketName,a=r.planId;if("string"!=typeof n)throw new Error("properties[BucketName] is not of type string");if("string"!=typeof a)throw new Error("properties[PlanId] is not of type string");var p={name:n,planId:a};i(e+"/planner/buckets",JSON.stringify(p),"POST",(function(e){"function"==typeof o&&o(e)}))}(0,r,(function(e){postResult({bucketId:e.id})}))}(0,o);break;default:throw new Error("The method "+t+" is not supported..")}}(r,0,n);break;default:throw new Error("The object "+t+" is not supported.")}}}();
//# sourceMappingURL=index.js.map
