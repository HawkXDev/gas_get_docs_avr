var app = SpreadsheetApp;
var ss = app.getActiveSpreadsheet();
var ws = ss.getActiveSheet();

function initMenu() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("My Menu");
    //menu.addItem("Go", "goFunc");
    //menu.addItem("Go 2", "goFunc2");
    //menu.addItem("AVR SPEC to XLSX", "avrSpecToXlsx");
    //menu.addItem("ZIPS", "zipsCreate");
    //menu.addItem("ZIPS COPY", "zipsCopy");
    menu.addItem("Discription", "discription");
    menu.addToUi();
}

function onOpen() {
    initMenu();
}

var dict = {};

function goFunc() {
    let folder = DriveApp.getFolderById("1ofEZ7Q1n-nPdQlMXzWvtgiLyBl98HbSA");
    let childFolder = folder.getFolders();

    while (childFolder.hasNext()) {
        var child = childFolder.next();

        let fldName = child.getName();

        var files = child.getFiles();
        var list = [];

        while (files.hasNext()) {
            file = files.next();
            //file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

            var row = [];
            row.push(file.getName(), file.getUrl());
            list.push(row);
        }

        dict[fldName] = list;
    }

    let lr = ws.getLastRow();

    for (let i = 1; i <= lr; i++) {
        let article = ws.getRange(i, 2).getValue();
        let brand = ws.getRange(i, 4).getValue();
        let typeAparat = ws.getRange(i, 5).getValue();

        let str = article + "\n" + brand + "\n" + typeAparat + "\n------\n";
        str += strFromDic(article);

        var rez = showInputBox(str);

        ws.getRange(i, 14).setValue(dict[article][rez][1]);
    }
}

function strFromDic(article) {
    let str = "";
    let fls = dict[article];

    for (let index = 0; index < fls.length; index++) {
        str += index + ". " + fls[index][0] + "\n";
    }

    return str;
}

function showInputBox(str) {
    let ui = SpreadsheetApp.getUi();
    let input = ui.prompt(str);
    return input.getResponseText();
}

function goFunc2() {
    let folder = DriveApp.getFolderById("1ofEZ7Q1n-nPdQlMXzWvtgiLyBl98HbSA");
    let childFolder = folder.getFolders();

    while (childFolder.hasNext()) {
        var child = childFolder.next();

        let fldName = child.getName();

        var files = child.getFiles();
        var list = [];

        while (files.hasNext()) {
            file = files.next();
            //file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

            var row = [];
            row.push(file.getName(), file.getUrl());
            list.push(row);
        }

        dict[fldName] = list;
    }

    let lr = ws.getLastRow();

    for (let i = 1; i <= lr; i++) {
        let article = ws.getRange(i, 1).getValue();

        let str = article + "\n------\n";
        str += strFromDic(article);

        var rez = showInputBox(str);

        ws.getRange(i, 2).setValue(dict[article][rez][1]);
    }
}

function avrSpecToXlsx() {
    let folderSpecXlsx = DriveApp.getFolderById("1rbxl6nwHKzKada3Va1t_C6i89yhgEXTc");

    let lr = ws.getLastRow();

    for (let i = 237; i <= 237; i++) {
        let article = ws.getRange(i, 1).getValue();
        let spec = ws.getRange("L" + i).getValue();

        let file = DriveApp.getFileById(getIdFromUrl(spec));
        let name = file.getName();

        var params = {
            method: "get",
            headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
        };

        var blob = UrlFetchApp.fetch(spec, params).getBlob();
        blob.setName(name + ".xlsx");

        var xlsxFile = folderSpecXlsx.createFile(blob);

        ws.getRange("P" + i).setValue(xlsxFile.getUrl());
    }
}

function zipsCreate(){
    let folderZip = DriveApp.getFolderById("1GkVOdPC5gQUkBGcbt_OGabrxs1rQzGsA");

    let lr = ws.getLastRow();

    for (let i = 2; i <= lr; i++) {
        let article = ws.getRange(i, 2).getValue();

        let pdfAvr = ws.getRange("M" + i).getValue();
        let pdfSchema = ws.getRange("N" + i).getValue();
        let zipXlsx = ws.getRange("P" + i).getValue();

        let filePdfAvr = DriveApp.getFileById(getIdFromUrl(pdfAvr));
        let filePdfSchema = DriveApp.getFileById(getIdFromUrl(pdfSchema));
        let fileZipXlsx = DriveApp.getFileById(getIdFromUrl(zipXlsx));

        fileZipXlsx.setName(fileZipXlsx.getName().split("/").join("_"));
        
        var blobs = [];
        blobs.push(filePdfAvr.getBlob());
        blobs.push(filePdfSchema.getBlob());
        blobs.push(fileZipXlsx.getBlob());

        var zipName = DriveApp.getFileById(getIdFromUrl(ws.getRange("K" + i).getValue())).getName()
            .replace("_spec", "")
            .split("/").join("_");
        zipName += ".zip";

        var zipBlob = Utilities.zip(blobs, zipName);

        var zipFile = folderZip.createFile(zipBlob);
        var zipFileId = zipFile.getId();
        var zipFileUrl = zipFile.getUrl();

        var douwnloadUrl = "https://drive.google.com/uc?export=download&id=" + zipFileId;

        zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        ws.getRange("O" + i).setValue(douwnloadUrl);
    }
}

function zipsCopy(){
    let folderAVRInfo= DriveApp.getFolderById("1fVTpsTgpnrtT4cUGNQuJcXtK0vxVyGHY");

    let lr = ws.getLastRow();

    for (let i = 2; i <= lr; i++) {
        let article = ws.getRange(i, 2).getValue();

        let zipUrl = ws.getRange("O" + i).getValue();
        let fileZip = DriveApp.getFileById(getIdFromUrl(zipUrl));

        var destfolder = folderAVRInfo.addFile(fileZip);
    }
}



function myFunction() {
  var folderId = "###"; // Please set the folder ID here.

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var blobs = [];
  while (files.hasNext()) {
    blobs.push(files.next().getBlob());
  }
  var zipBlob = Utilities.zip(blobs, folder.getName() + ".zip");
  var fileId = DriveApp.createFile(zipBlob).getId();
  var url = "https://drive.google.com/uc?export=download&id=" + fileId;
  Logger.log(url);
}
function getSubFolders(parent) {
    parent = parent.getId();
    var childFolder = DriveApp.getFolderById(parent).getFolders();
    while (childFolder.hasNext()) {
        var child = childFolder.next();
        Logger.log(child.getName());
        getSubFolders(child);
    }
    return;
}
function listFolders(folderId) {
    var parentFolder = DriveApp.getFolderById(folderId);
    var childFolders = parentFolder.getFolders();
    while (childFolders.hasNext()) {
        var child = childFolders.next();
        Logger.log(child.getName());
        getSubFolders(child);
    }
}
function list_all_files_inside_one_folder_without_subfolders() {
    var sh = SpreadsheetApp.getActiveSheet();
    var folder = DriveApp.getFolderById('0B3qSFd3iikE3TERRSHExa29SU3M'); // I change the folder ID  here 
    var list = [];
    list.push(['Name', 'ID', 'Size']);
    var files = folder.getFiles();
    while (files.hasNext()) {
        file = files.next();
        var row = []
        row.push(file.getName(), file.getId(), file.getSize())
        list.push(row);
    }
    sh.getRange(1, 1, list.length, list[0].length).setValues(list);
}
function convertSheetToXLSX() {
    var sheetId = "2SqIXLiic6-gjI2KwQ6OIgb-erbl3xqzohRgE06bfj2c";
    var spreadsheetName = "My Spreadsheet";
    var destination = DriveApp.getFolderById("1vFL98cgKdMHLNLSc542pUt4FMRTthUvL");
    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + sheetId + "&exportFormat=xlsx";
    var params = {
        method: "get",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
    };
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(spreadsheetName + ".xlsx");
    destination.createFile(blob);
}
function getIdFromUrl(url) {
    return url.match(/[-\w]{25,}/); 
}