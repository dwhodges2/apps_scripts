// Script to merge rows based on column 0 and aggregate data from cols 1, 2, 3 for each.
// Example: 
// a,1,2,3
// b,1,2,3
// a,1,2,3
// becomes:
// a,2,4,6
// b,1,2,3

function removeDuplicates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var newData = new Array();
    var cumulative = {};
    var dupeRow = {};

    for (i in data) {
        var row = data[i];
        var duplicate = false;
        for (j in newData) {
            if (row[0] == newData[j][0]) {
                duplicate = true;
                dupeRow = j
            }
        }

        if (!duplicate) {
            newData.push(row);
        }

        // Adjust below for number of columns to aggregate.
        else {
            var theCells = new Array();
            theCells[1] = sumData(row[1], newData[dupeRow][1]);
            theCells[2] = sumData(row[2], newData[dupeRow][2]);
            theCells[3] = sumData(row[3], newData[dupeRow][3]);
            //  theCells[4] = sumData(row[4], newData[dupeRow][4]);
            newData[dupeRow].splice(1, 3, theCells[1], theCells[2], theCells[3]);
        }

    }

    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}


function sumData(oldVal, newVal) {
    var x;
    x = oldVal + newVal;

    return x;
}
