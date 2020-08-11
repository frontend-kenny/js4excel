const xlsx = require('node-xlsx') //refer to https://github.com/mgcrea/node-xlsx
const fs = require("fs")
const CONST = require('./src/const')

let extractEmailAddrPerCat = function (emailItem) {
    let emailItemBody = emailItem[CONST.INDEX_OF_COLUMN_IN_INPUT_BODY]
    let emailItemCategory = emailItem[CONST.INDEX_OF_COLUMN_IN_INPUT_CATEGORIES]
    switch (emailItemCategory) {
        case CONST.CAM2:
            {
                result = /RCPT TO:<(.*?)>:/.exec(emailItemBody)
                // console.log('result: ', result)
                if (result && result.length > 1) {
                    return result[1]
                }
      /*           let reg = /RCPT TO:<([(\w\W)]*)>:/
                emailItemBodyArray = emailItemBody.split('\n')
                for (let i = 0; i < emailItemBodyArray.length; i++) {
                    result = reg.exec(emailItemBodyArray[i])
                    console.log('result: ', result)
                    if (result && result.length > 1) {
                        return result[1]
                    }
                } */
            }
        case CONST.CAM3: {

        }
        case CONST.CAM4: {

        }
        case CONST.CAM6: {

        }
        case CONST.CAM7: {

        }
        default: {
            return 'extracted failed'
        }
    }
}

let createOutputDataArrayItem = function (emailItem, extractedEmailAddress) {

    let emailItemCategory = emailItem[CONST.INDEX_OF_COLUMN_IN_INPUT_CATEGORIES]
    let emailItemFromName = emailItem[CONST.INDEX_OF_COLUMN_IN_INPUT_FROM_NAME]

    let outputDataArrayItem = []
    outputDataArrayItem[CONST.OUTPUT_COLUMN_NUMBER_0] = emailItemFromName
    outputDataArrayItem[CONST.OUTPUT_COLUMN_NUMBER_1] = emailItemFromName
    outputDataArrayItem[CONST.OUTPUT_COLUMN_NUMBER_2] = emailItemCategory
    outputDataArrayItem[CONST.OUTPUT_COLUMN_NUMBER_3] = extractedEmailAddress
    outputDataArrayItem[CONST.OUTPUT_COLUMN_NUMBER_4] = extractedEmailAddress

    return outputDataArrayItem
}

let writeFileFromButter = function (outputFileName, outputDataArray) {
    //Custom column width
    const options = { '!cols': [{ wch: 30 }, { wch: 30 }, { wch: 6 }, { wch: 30 }, { wch: 30 }] };
    //building a xlsx,return a buffer
    let buffer = xlsx.build([{ name: "CBN-output", data: outputDataArray }], options);
    //write to file
    fs.writeFile(outputFileName, buffer, function (err) {
        if (err) {
            return console.error(err);
        }
        console.log("数据写入成功！");
    })
}

// vars
let inputFileName = 'CBN input.CSV'
let outputFileName = 'CBN output.xlsx'
let outputDataArray = []
let emailItem = []
let extractedEmailAddress = '?'
let inputFile = `${__dirname}/file/${inputFileName}`
let outputFile = `${__dirname}/file/${outputFileName}`

//1 Parsing a xlsx from file/buffer, outputs an array of worksheets
// const workSheets = xlsx.parse(fs.readFileSync(inputFileName)); // Parse a buffer
const workSheets = xlsx.parse(inputFile) // Parse a file

//2 Get the data from the first sheet
const workSheetsData = workSheets[CONST.ZERO].data


//4 Loop every email,skip the line of subject,index from 1
for (let i = 0; i < workSheetsData.length; i++) {
    if (i === 0) {
        // const subjectArray =  workSheetsData[i]
    } else {
        //4.1 Get each email as array
        emailItem = workSheetsData[i]
        //4.2 Find out the from-email per Categories!!!!!
        extractedEmailAddress = extractEmailAddrPerCat(emailItem)
        //4.3 fill in the outputDataArray
        outputDataArray[i - 1] = createOutputDataArrayItem(emailItem, extractedEmailAddress)
    }
}
//5 write the output-file from output data array
writeFileFromButter(outputFile, outputDataArray)
console.log('outputDataArray: ', outputDataArray)
