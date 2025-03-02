import { google } from 'googleapis';
import dotenv from 'dotenv';

dotenv.config();

const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets"
});

const service = google.sheets({ version: 'v4', auth });
const spreadsheetId = process.env.SPREADSHEET_ID;

async function updateValues(range, valueInputOption, resource) {

    const result = await service.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        resource
    });
    console.log('%d cells updated.', result.data.updatedCells);
    return result;
}

async function batchUpdateValues(resource) {

    const result = await service.spreadsheets.values.batchUpdate({
        spreadsheetId,
        resource
    });
    console.log('%d cells updated.', result.data.updatedCells);
    return result;
}

async function appendValues(range, valueInputOption, resource) {

    const result = await service.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption,
        insertDataOption: 'INSERT_ROWS',
        resource
    });
    console.log('%d cells updated.', result.data.updatedCells);
    return result;
}

async function getValues(range) {

    const result = await service.spreadsheets.values.get({
        spreadsheetId,
        range,
        majorDimension:"ROWS", //Optional, Default: ROWS (We can use COLUMNS as well)
        valueRenderOption: "FORMATTED_VALUE" //Optional, Default: FORMATTED_VALUE, other options are UNFORMATTED_VALUE, FORMULA
    });
    console.log("Values obtained from sheet: ",result.data.values);
}

async function batchGetValues(ranges) {

    const result = await service.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges
    });
    result.data.valueRanges.forEach(valueRange => {
        console.log(valueRange.values);
    })
}

async function updateBatch(requests) {
    await service.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
            requests
        }
    });
}

//update and append values starts ===============================
let updateAndAppendValuesResource = {
    // majorDimension:"COLUMNS",
    values: [
        ["Team Name", "Number of team members"],
        ["eStatement", 10],
        ["PD", 20]
    ]
};
updateValues("TeamInfo!A6", "USER_ENTERED", updateAndAppendValuesResource);

appendValues("TeamInfo!B4", "USER_ENTERED", updateAndAppendValuesResource);

//update and append values ends ===============================


//update batch values starts ===============================
const data = [
    {
        range:"TeamInfo!D1:E3",
        values: [
            ["Location","Office Time"],
            ["House: 185", "12PM - 8PM"],
            ["House: 185", "12PM - 8PM"]
        ]
    },
    {
        range:"TeamInfo!G1:G3",
        values: [
            ["Type"],
            ["Onshore-Offshore"],
            ["Onshore-Offshore"]
        ]
    }
]

const resource = {
    data,
    valueInputOption:"USER_ENTERED"
}
batchUpdateValues(resource);
//update batch values ends ===============================

//get values starts ===============================
getValues("TeamInfo!A1:B3");
//get values ends ===============================

//get batch values starts ===============================
const ranges = [
    "TeamInfo!A1:B1",
    "TeamInfo!A3:B3"
];

batchGetValues(ranges);
//get batch values ends ===============================

//update batch starts =========================
const requests = []

requests.push({
    repeatCell: {
        range: {
            sheetId: 0,
            startRowIndex: 0,
            endRowIndex: 3,
            startColumnIndex: 0,
            endColumnIndex: 7
        },
        cell: {
            userEnteredFormat: {
                // wrapStrategy: "WRAP",
                backgroundColor: {
                    red: 1,
                    green: 0.949,
                    blue: 0.8
                },
                textFormat:{
                    bold: true,
                },
                horizontalAlignment: "CENTER",
                verticalAlignment: "TOP"
            } 
        },
        fields: "userEnteredFormat.wrapStrategy,userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
    }
});

requests.push({
    autoResizeDimensions: {
        dimensions: {
        sheetId: 0,
        dimension: 'COLUMNS', 
        startIndex: 0,
        endIndex: 7,
        }
    }
});

updateBatch(requests);
//update batch ends =========================

//conditional formatting starts =========================
const conditionalFormattingRequests = [];

conditionalFormattingRequests.push({
    addConditionalFormatRule: {
        rule: {
            ranges: [
                {
                    sheetId: 0,
                    startRowIndex: 1,
                    endRowIndex: 3,
                    startColumnIndex: 1,
                    endColumnIndex: 2
                }
            ],
            booleanRule: {
                condition: {
                    type: "NUMBER_GREATER",
                    values: [
                        {
                            userEnteredValue: "10"
                        }
                    ]
                },
                format: {
                    textFormat: {
                        italic: true
                    },
                    backgroundColor: {
                        red: 1,
                        green: 0.5,
                        blue: 0
                    }
                }
            }
        },
        index: 0
    }
});

updateBatch(conditionalFormattingRequests);
//conditional formatting ends =========================