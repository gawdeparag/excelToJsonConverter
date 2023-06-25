const xlsx = require('xlsx');

var fs = require('fs');

const file = xlsx.readFile('Json Table -3.xlsx');

// Grab the sheet info from the file
const sheetNames = file.SheetNames;
const totalSheets = sheetNames.length;

// Variable to store our data
let parsedData = [];

function convertExcelToJson(sheetNames, totalSheets) {

    // Loop through sheets
    for (let i = 0; i < totalSheets; i++) {

        // Convert to json using xlsx
        const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]]);

        // Add the sheet's json to our data array
        parsedData.push(...tempData);
    }
    return parsedData;
}

function generateJSONFile(data) {
    try {
        fs.writeFileSync('data.json', JSON.stringify(data))
    } catch (err) {
        console.error(err)
    }
}

convertExcelToJson(sheetNames, totalSheets);

// call a function to save the data in a json file
generateJSONFile(parsedData);