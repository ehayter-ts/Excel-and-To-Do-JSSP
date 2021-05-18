!function(){metadata={systemName:"MSExcelTodoJSSP",displayName:"Microsoft Excel and To Do",description:"A connector for Microsoft Excel and To Do"};const e="https://graph.microsoft.com/v1.0";function t(e,t,i,r){var o=new XMLHttpRequest;o.onreadystatechange=function(){if(console.log("ExecuteRequest XHR status: "+o.status+","+o.responseText),4===o.readyState)if(201==o.status){try{e=JSON.parse(o.responseText)}catch(e){}"function"==typeof r&&r(e)}else if(204==o.status)"function"==typeof r&&r(o.responseText);else if(200==o.status){var e=JSON.parse(o.responseText);"function"==typeof r&&r(e)}else if(202==o.status)"function"==typeof r&&r(null);else{if(400==o.status){e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if(404!=o.status){postResult({});e=JSON.parse(o.responseText);throw new Error(e.error.code+": "+e.error.message)}if("No Team found with Group id"!=(e=JSON.parse(o.responseText)).error.message.startswith)throw new Error(e.error.code+": "+e.error.message)}},console.log("MSTeamsConnector ExecuteRequest: "+e),o.open(i.toUpperCase(),e),o.withCredentials=!0,o.setRequestHeader("Accept","application/json"),"PUT"!=i.toUpperCase()&&"POST"!=i.toUpperCase()&&"PATCH"!=i.toUpperCase()||o.setRequestHeader("Content-Type","application/json"),o.send(t)}ondescribe=function(){postSchema({objects:{drive:{displayName:"Drive",description:"Drive",properties:{fileId:{displayName:"File ID",description:"File ID",type:"string"},fileWebUrl:{displayName:"File Web URL",description:"File Web URL",type:"string"},fileSize:{displayName:"File Size",description:"File Size",type:"number"},fileName:{displayName:"File Name",description:"File Name",type:"string"},fileCreatedDate:{displayName:"File Created Date",description:"File Created Date",type:"dateTime"},fileCreatedBy:{displayName:"File Created By",description:"File Created By",type:"string"},fileCreatedByEmail:{displayName:"File Created By Email",description:"File Created By Email",type:"string"},fileMimeType:{displayName:"File Mime Type",description:"File Mime Type",type:"string"},filePath:{displayName:"File Path",description:"File Path",type:"string"},folderName:{displayName:"Folder Name",description:"Folder Name",type:"string"},folderPath:{displayName:"Folder Path",description:"Folder Path",type:"string"}},methods:{searchFile:{displayName:"Search for a file in my OneDrive",type:"read",inputs:["fileName","filePath"],requiredInputs:["fileName"],outputs:["fileId","fileWebUrl","fileSize","fileName","fileCreatedDate","fileCreatedBy","fileCreatedByEmail","fileMimeType"]},createFolder:{displayName:"Create Folder",type:"execute",inputs:["folderName","folderPath"],requiredInputs:["folderName"],outputs:[]}}},excel:{displayName:"Excel",description:"Excel",properties:{fileId:{displayName:"File ID",description:"File ID",type:"string"},sheetName:{displayName:"Sheet Name",description:"Sheet Name",type:"string"},column1:{displayName:"Column 1",description:"Column 1",type:"string"},column2:{displayName:"Column 2",description:"Column 2",type:"string"},column3:{displayName:"Column 3",description:"Column 3",type:"string"},column4:{displayName:"Column 4",description:"Column 4",type:"string"},column5:{displayName:"Column 5",description:"Column 5",type:"string"},column6:{displayName:"Column 6",description:"Column 6",type:"string"},column7:{displayName:"Column 7",description:"Column 7",type:"string"},column8:{displayName:"Column 8",description:"Column 8",type:"string"},column9:{displayName:"Column 9",description:"Column 9",type:"string"},column10:{displayName:"Column 10",description:"Column 10",type:"string"},column11:{displayName:"Column 11",description:"Column 11",type:"string"},column12:{displayName:"Column 12",description:"Column 12",type:"string"},column13:{displayName:"Column 13",description:"Column 13",type:"string"},column14:{displayName:"Column 14",description:"Column 14",type:"string"},column15:{displayName:"Column 15",description:"Column 15",type:"string"},column16:{displayName:"Column 16",description:"Column 16",type:"string"},column17:{displayName:"Column 17",description:"Column 17",type:"string"},column18:{displayName:"Column 18",description:"Column 18",type:"string"},column19:{displayName:"Column 19",description:"Column 19",type:"string"},column20:{displayName:"Column 20",description:"Column 20",type:"string"}},methods:{getUsedRangeItems:{displayName:"Get Worksheet Rows in Used Range",type:"list",inputs:["fileId","sheetName"],requiredInputs:["fileId","sheetName"],outputs:[]}}}}})},onexecute=function({objectName:i,methodName:r,parameters:o,properties:s}){switch(i){case"drive":!function(i,r,o){switch(i){case"searchFile":!function(e,i){!function(e,i,r){let o=i.fileName,s=i.filePath;if("string"!=typeof o)throw new Error("properties[FileName] is not of type string");var n;n="string"==typeof s&&""!=s?`https://graph.microsoft.com/v1.0/me/drive/root:/${s}:/search(q='${o}')`:`https://graph.microsoft.com/v1.0/me/drive/root/search(q='${o}')`;t(n,null,"GET",(function(e){"function"==typeof r&&r(e)}))}(0,i,(function(e){var t={fileId:"",fileWebUrl:"",fileSize:0,fileName:"",fileCreatedDate:null,fileCreatedBy:"",fileCreatedByEmail:"",fileMimeType:""};e.value.length>0&&(t.fileId=e.value[0].id,t.fileWebUrl=e.value[0].webUrl,t.fileSize=e.value[0].size,t.fileName=e.value[0].name,t.fileCreatedDate=Date.parse(e.value[0].createdDateTime),t.fileCreatedBy=e.value[0].createdBy.user.displayName,t.fileCreatedByEmail=e.value[0].createdBy.user.email,t.fileMimeType=e.value[0].file.mimeType),postResult(t)}))}(0,o);break;case"createFolder":!function(i,r){!function(i,r,o){let s=r.folderName,n=r.folderPath;if("string"!=typeof s)throw new Error("properties[FolderName] is not of type string");var a;a="string"==typeof n&&""!=n?`https://graph.microsoft.com/v1.0/me/drive/root:/${n}:/children'`:e+"/me/drive/root/children";var l={name:s,folder:{},"@microsoft.graph.conflictBehavior":"replace"};t(a,JSON.stringify(l),"POST",(function(){"function"==typeof o&&o()}))}(0,r,(function(){}))}(0,o);break;default:throw new Error("The method "+i+" is not supported..")}}(r,0,s);break;case"excel":!function(e,i,r){switch(e){case"getUsedRangeItems":!function(e,i){!function(e,i,r){let o=i.fileId,s=i.sheetName;if("string"!=typeof o)throw new Error("properties[FileId] is not of type string");if("string"!=typeof s)throw new Error("properties[ExcelSheetName] is not of type string");t(`https://graph.microsoft.com/v1.0/me/drive/items/${o}/workbook/worksheets('${s}')/usedRange`,null,"GET",(function(e){"function"==typeof r&&r(e)}))}(0,i,(function(e){var t=e.text.map(e=>{for(var t={},i=0;i<e.length;i++)i+1<21&&(t["column"+(i+1)]=e[i]);return t});postResult(t)}))}(0,r);break;default:throw new Error("The method "+e+" is not supported..")}}(r,0,s);break;default:throw new Error("The object "+i+" is not supported.")}}}();
//# sourceMappingURL=index.js.map
