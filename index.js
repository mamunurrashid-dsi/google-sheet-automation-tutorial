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


// let resource = {
//     // majorDimension:"COLUMNS",
//     values: [
//         ["Team Name", "Number of team members"],
//         ["eStatement", 10],
//         ["PD", 20]
//     ]
// };
// updateValues("TeamInfo!A6", "USER_ENTERED", resource);

let resource = {
    values: [
        ["Team Name", "Number of team members"],
        ["eStatement", 10],
        ["PD", 20]
    ]
};
appendValues("TeamInfo!B4", "USER_ENTERED", resource);

// const data = [
//     {
//         range:"TeamInfo!D1:E3",
//         values: [
//             ["Location","Office Time"],
//             ["House: 185", "12PM - 8PM"],
//             ["House: 185", "12PM - 8PM"]
//         ]
//     },
//     {
//         range:"TeamInfo!G1:G3",
//         values: [
//             ["Type"],
//             ["Onshore-Offshore"],
//             ["Onshore-Offshore"]
//         ]
//     }
// ]

// const resource = {
//     data,
//     valueInputOption:"USER_ENTERED"
// }
// batchUpdateValues(resource);

// getValues("TeamInfo!A1:B3");

// const ranges = [
//     "TeamInfo!A1:B1",
//     "TeamInfo!A3:B3"
// ];

// batchGetValues(ranges);