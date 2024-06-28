import xlsx from 'xlsx';
import axios from 'axios'
import { uploadFile, changeData, deleteFile, generatePublicURL } from './weniDrive.mjs'

import * as fs from 'fs'


function donwloadSheets(sheetID, sheetName) {

    const donwloadResponse = axios.get(`https://docs.google.com/spreadsheets/d/${sheetID}/export?format=xlsx`, {
        responseType:'stream'
    })
    .then((data) => {
        data.data.pipe(fs.createWriteStream(`${sheetName}.xlsx`))
        return true
    })
    .catch((err) => {
        console.log(err.toString())
        return new Error(err.toString())
    })
}

donwloadSheets("15y8rf1w6p4XCaDH-YkS4yJYAby_AkcMfrp1j1UeyAK0", "tutores_sheet_id")

//Essa função pode ser reutilizada
/* function readTutoresSheet() {

    try {
        const tutorNames = []    
        const workBook = xlsx.readFile('tutoresExcel.xlsx')
        let workSheet = workBook.Sheets[workBook.SheetNames[0]]

        for(let i = 2; i <= 100; i++) {
            if(workSheet[`F${i}`] == undefined) {
                break;
            }
            tutorNames.push(workSheet[`F${i}`].v)
        }

        const tutorNameAndSurname = tutorNames.map((value) => {
            const splittedValue = value.split(" ")
            return splittedValue[0] + " " + splittedValue[1]
        })

        readStudentsByGroup(tutorNameAndSurname)
    } catch(err) {
        console.log(err) 
        throw new Error(err.toString())
    }
} */

function formatTutorName(tutor) {

    const tutorNameAndSurname = []

    if(tutor.split(" ").length < 2) {
        throw new Error("You need to provide your name and surname")
    }

    const firstName = tutor.split(" ")[0].charAt(0).toUpperCase() + tutor.split(" ")[0].slice(1)
    const surname = tutor.split(" ")[1].charAt(0).toUpperCase() + tutor.split(" ")[1].slice(1)
    const nameAndSurname = firstName + " " + surname
    tutorNameAndSurname.push(nameAndSurname)

    readStudentsByGroup(tutorNameAndSurname)
}

async function readStudentsByGroup(listaTutores) {

    for(let i = 0; i <= listaTutores.length - 1; i++) {

        let groupUUID = ""

        const groupResponse = await axios.get(`https://flows.weni.ai/api/v2/groups.json`, {
            headers: {
                Authorization: `Token 01606573d6244a492320e2fa2231f369f6c8d0b1`
            },
            params: {
                name: `Tutor: ${listaTutores[i]}`
            }
        })
        .then((data) => {
            if(data.data.results.length == 0) {
                console.log(`Nenhum grupo encontrado para esse professor: ${listaTutores[i]}`)
                return false
            } else {
                groupUUID = data.data.results[0].uuid
                return true
            }
        })
        .catch((err) => {
            console.log(err.toString())
            return false
        })
        
        if(groupResponse) {
            getStudentsByTutor(listaTutores[i], groupUUID) 
        }

    }

}

async function getStudentsByTutor(tutor, groupUUID) {

    try {
        const studentList = []

        const studentResponse = await axios.get(`https://flows.weni.ai/api/v2/contacts.json`, {
            headers: {
                Authorization: `Token 01606573d6244a492320e2fa2231f369f6c8d0b1`
            },
            params: {
                group: groupUUID
            }
        })
        .then((data) => {
            for(let i = 0; i <= data.data.results.length - 1; i++) {
                studentList.push(
                    {   
                        name: data.data.results[i].name,
                        studentID: data.data.results[i].uuid,
                        age: data.data.results[i].fields.age,
                        pin: data.data.results[i].fields.access_pin,
                        username: data.data.results[i].urns[0],
                    }
                )
            }
        })
        .catch((err) => {
            console.log(err.toString())
        })

        createExcel(tutor, studentList)

    }catch(err) {
        console.log(err.toString())
    }

}

async function createExcel(tutor, studentList) {

    const sheetList = studentList.map((element) => {
        return [tutor, element.name, element.studentID, element.age, element.pin, element.username]
    })

    
    const data = [
        ["TUTOR", "STUDENT_NAME", "STUDENT_UUID", "STUDENT_AGE", "STUDENT_PIN", "STUDENT_USERNAME"],
        ...sheetList
    ]

    const workBook = xlsx.utils.book_new();
    workBook.SheetNames.push("StudentList");
    workBook.Sheets["StudentList"] = xlsx.utils.aoa_to_sheet(data)
    xlsx.writeFile(workBook, `./sheets/${tutor.replace(" ", "_")}.xlsx`)

    const fileID = await uploadFile(tutor.replace(" ", "_"))
    updateTutoresSheetID(tutor, fileID)
    
}

async function updateTutoresSheetID(tutor, id) {

    try {
        const workBook = xlsx.readFile('tutores_sheet_id.xlsx')
        let workSheet = workBook.Sheets[workBook.SheetNames[0]]
        let oldDocumentID = ""
        let sheetDonwloadURL = ""

        for(let i = 2; i <= 100; i++) {
            if(workSheet[`A${i}`] == undefined) {
                break;
            } else if(workSheet[`A${i}`].v === tutor) {
                oldDocumentID = workSheet[`B${i}`].v
                deleteFile(oldDocumentID)
                sheetDonwloadURL = await generatePublicURL(id)
                changeData(`B${i}`, id)
                changeData(`C${i}`, sheetDonwloadURL)
                return sheetDonwloadURL
            }
        }

        sheetDonwloadURL = await generatePublicURL(id)

        axios.get(`https://script.google.com/macros/s/AKfycbyHQ26luadyjaCp9rKnR0K5XjND-tXVHsZPTUnGwm1cEZTxTfxwTGar31MIiw1Yiv0qIw/exec?gid=0`, {
            params: {
                "TUTOR": tutor,
                "SHEET_ID": id,
                "SHEET_DONWLOAD_URL": sheetDonwloadURL
            }
        })

        console.log(sheetDonwloadURL)
        return sheetDonwloadURL

    }catch(err) {
        console.log(err)
    }

}

formatTutorName("alicia marcela")