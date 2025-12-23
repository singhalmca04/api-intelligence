const { PDFParse } = require('pdf-parse');
const questions = require('./questions');
const excelHelper = require('../helpers/excel');
const path = require('path');
const fs = require('fs');
//const mastery = require('../helpers/updated_mastery.xlsm')
async function uploadPdf(req, res) {
    try {
        let marks = [];
        let myurl = 'http://localhost:3000/uploads/' + req.file.filename;
        const parser = new PDFParse({ url: myurl });

        const result = await parser.getText();
        let fileData = result.text.split("\n");
        // Extract name, dob, etc.
        const profileData = excelHelper.extractProfileData(fileData);

        console.log("Name:", profileData.name);
        console.log("DOB:", profileData.dob);
        console.log("Email:", profileData.email);

        for (let i = 0; i < questions.length; i++) {
            let index = fileData.findIndex((data) => data === questions[i]);
            if (index >= 0) {
                if (fileData[index + 1] === 'Mostly Disagree') {
                    marks.push(1);
                } else if (fileData[index + 1] === 'Slightly Disagree') {
                    marks.push(2)
                } else if (fileData[index + 1] === 'Slightly Agree') {
                    marks.push(3)
                } else {
                    marks.push(4)
                }
            } else {
                let halfOfQuestion = questions[i].substring(0, questions[i].length / 2);

                let matchingQuestion = fileData.filter(data => data.includes(halfOfQuestion));

                let index1 = fileData.findIndex((data) => data === matchingQuestion[0]);
                if (fileData[index1 + 2] === 'Mostly Disagree') {
                    marks.push(1);
                } else if (fileData[index1 + 2] === 'Slightly Disagree') {
                    marks.push(2)
                } else if (fileData[index1 + 2] === 'Slightly Agree') {
                    marks.push(3)
                } else {
                    marks.push(4)
                }
            }
        }

        await excelHelper.updateMarksColumn(
            "./helpers/mastery.xlsm", // file path
            "Sheet1",                 // sheet tab name
            6,                        // starting row number
            "C",                      // marks column letter
            marks,                     // marks array
            profileData
        );
        // const filePath = path.join(__dirname, 'upadted_mastery.xlsm');
        //const filePath = path.join('../helpers/upadted_mastery.xlsm');
        const filePath = path.join(__dirname, '..', 'helpers', 'updated_mastery.xlsm');
        console.log(filePath, 'filepath')
        res.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12');
        res.setHeader('Content-Disposition', 'attachment; filename=' + profileData.name+'.xlsm');

        // Send the file
        res.status(200).sendFile(filePath);
        // if (fs.existsSync(filePath)) {
        //     // Set headers for file download
        //     res.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12');
        //     res.setHeader('Content-Disposition', 'attachment; filename=' + 'your_file.xlsm');

        //     // Send the file
        //     res.sattus(200).sendFile(filePath);
        // } else {
        //     res.status(404).send('File not found');
        // }
        //res.status(200).send('PDF processed successfully.');
    } catch (err) {
        console.log(err);
        res.status(500).send('Error processing PDF.');
    }
}
module.exports = {
    uploadPdf
}