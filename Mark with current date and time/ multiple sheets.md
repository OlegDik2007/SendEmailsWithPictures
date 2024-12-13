




function sendEmailsWithPictures_test1() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'Test1'; // Replace 'Sheet1' with the name of your desired sheet
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
        throw new Error(`Sheet with name '${sheetName}' not found.`);
    }

    const data = sheet.getDataRange().getValues();
    const folderId = '1IdTT3Dyub3UfY7Dhawru8t1Y-enFkO8A'; // Replace with your actual folder ID
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    const fileMap = {};
    while (files.hasNext()) {
        const file = files.next();
        fileMap[file.getName()] = file.getId();
    }

    for (let i = 1; i < data.length; i++) { // Start from row 2 if row 1 is headers
        const email = data[i][0];
        const subject = data[i][1];
        const message = data[i][2];
        const fileName = data[i][3];

        if (fileMap[fileName]) {
            const fileId = fileMap[fileName];
            const attachment = DriveApp.getFileById(fileId).getBlob();
            GmailApp.sendEmail(email, subject, message, {
                attachments: [attachment]
            });
            sheet.getRange(i + 1, 5).setValue(new Date()); // Mark with current date and time
        } else {
            sheet.getRange(i + 1, 5).setValue("File Not Found");
        }
    }
}
