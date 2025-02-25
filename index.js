import { google } from 'googleapis';
import dotenv from 'dotenv';

dotenv.config();

const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets"
});

const service = google.sheets({ version: 'v4', auth });

async function insertValues(spreadsheetId, range, valueInputOption, resource) {

    const result = await service.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        resource
    });
    console.log('%d cells updated.', result.data.updatedCells);
    return result;
}

let resource = {
    values: [
        ["Team Name", "Number of team members"],
        ["eStatement", 10],
        ["PD", 20]
    ]
};
console.log("spreadsheet id: ",process.env.SPREADSHEET_ID);
insertValues(process.env.SPREADSHEET_ID, "TeamInfo!A1:B3", "USER_ENTERED", resource);