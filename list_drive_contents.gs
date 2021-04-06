// Functions to populate sheet with list of contents of 
// given Drive folder and subfolders. Run from "Scripts" menu in document.

function onOpen() {
    // Set up script menu
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Scripts')
        .addItem('List folder contents (named folder)', 'StartListingByName')
        .addItem('List folder contents (folder id)', 'StartListingById')
        .addToUi();
}

function ListFolderContentsById(sheet, folder_name) {
    // Get folder pointer
    var folder = DriveApp.getFolderById(folder_name);
    var current_folder = folder;
    // Get files and links in folder
    var file_names = current_folder.getFiles();
    var file_name;
    var name;
    var link;
    while (file_names.hasNext()) {
        file_name = file_names.next();
        name = file_name.getName();
        link = file_name.getUrl();
        sheet.appendRow([name, link]);
    }
    // Get sub-folders in folder
    var sub_folder_names = current_folder.getFolders();
    var sub_folder_name;
    var name;
    while (sub_folder_names.hasNext()) {
        sub_folder_name = sub_folder_names.next();
        name = sub_folder_name.getName();
        sub_id = sub_folder_name.getId();
        sheet.appendRow([name]);
        ListFolderContentsById(sheet, sub_id);
    }
}

function ListFolderContentsByName(sheet, folder_name) {
    // Get folder pointer
    var folder = DriveApp.getFoldersByName(folder_name);
    var current_folder = folder.next();
    // Get files and links in folder
    var file_names = current_folder.getFiles();
    var file_name;
    var name;
    var link;
    while (file_names.hasNext()) {
        file_name = file_names.next();
        name = file_name.getName();
        link = file_name.getUrl();
        sheet.appendRow([name, link]);
    }
    // Get sub-folders in folder
    var sub_folder_names = current_folder.getFolders();
    var sub_folder_name;
    var name;
    while (sub_folder_names.hasNext()) {
        sub_folder_name = sub_folder_names.next();
        name = sub_folder_name.getName();
        sub_id = sub_folder_name.getId();
        sheet.appendRow([name]);
        ListFolderContentsById(sheet, sub_id);
    }
}

function StartListingByName() {
    var ui = SpreadsheetApp.getUi();
    var input = (ui.prompt('Folder name:'));
    folder_name = (input.getResponseText());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    sheet.clear();
    sheet.appendRow(['Name', 'Link']);
    ListFolderContentsByName(sheet, folder_name);
}

function StartListingById() {
    var ui = SpreadsheetApp.getUi();
    var input = (ui.prompt('Folder UID:'));
    folder_id = (input.getResponseText());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    sheet.clear();
    sheet.appendRow(['Name', 'Link']);
    ListFolderContentsById(sheet, folder_id);
}