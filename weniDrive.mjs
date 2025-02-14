import * as path from 'path'
import * as fs from 'fs'
import { google } from 'googleapis'

const clientID = "mock";
const clientSecret = "mock";

const redirectURI = "mock";
const refreshToken = "mock"


const oAuthClient = new google.auth.OAuth2(
    clientID,
    clientSecret,
    redirectURI
)

oAuthClient.setCredentials({refresh_token: refreshToken})

const drive = google.drive(
    {
        version: 'v3',
        auth: oAuthClient
    }
)

const sheets = google.sheets(
    {
        version: "v4",
        auth: oAuthClient
    }
)

export async function changeData(position, documentID) {

    const response = await sheets.spreadsheets.values.batchUpdate(
        {
            spreadsheetId: "15y8rf1w6p4XCaDH-YkS4yJYAby_AkcMfrp1j1UeyAK0",
            requestBody: {
                "valueInputOption": "RAW",
                "data": [
                    {
                        "range": position,
                        "majorDimension": "ROWS",
                        "values": [
                            [documentID]
                        ]
                    }
                ]
            }
        }
    )
}

export async function uploadFile(fileName) {

    const pathFile = path.join("./sheets/", `${fileName}.xlsx`)

    const fileData = {
        name: `${fileName}.xlsx`,
        parents: ['15xyTUhhS5tFeXli7oqxAo-QMqtWB31ws']
    }

    const media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(pathFile)
    }
    
    try {

        const response = await drive.files.create(
            {
                resource: fileData,
                media: media,
                enforceSingleParent: true,
                includeLabels: "tutores"
            }
        )

        return response.data.id

    } catch(err) {
        console.log(err.toString())
    }

    
}

export async function deleteFile(fileID) {

    try {
        const deleteResponse = await drive.files.delete(
            {
                fileId: fileID
            }
        )

        return deleteResponse
    }catch(err) {
        console.log(err.toString())
    }

}

export async function generatePublicURL(fileID) {

    try {
        const generationResponse = await drive.files.get(
            {
                fileId: fileID,
                fields: "webContentLink"
            }
        )

        return generationResponse.data.webContentLink
    }catch(err) {
        console.log(err.toString())
    }

}

generatePublicURL("1KejIGjSXS6BYN9VRcd3wNVqlp6wtOPEi")
