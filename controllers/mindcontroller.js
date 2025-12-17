const { PDFParse } = require('pdf-parse');
const questions = require('./questions');
const excelHelper = require('../helpers/excel');
async function uploadPdf(req, res) {
    try {
        let marks = [];
        let myurl = 'http://localhost:3000/uploads/' + req.file.filename;
        const parser = new PDFParse({ url: myurl });

        const result = await parser.getText();
        let fileData = result.text.split("\n");
        
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
                if (fileData[index + 2] === 'Mostly Disagree') {
                    marks.push(1);
                } else if (fileData[index + 2] === 'Slightly Disagree') {
                    marks.push(2)
                } else if (fileData[index + 2] === 'Slightly Agree') {
                    marks.push(3)
                } else {
                    marks.push(4)
                }
            }
        }
        
        excelHelper.updateMarksColumn(
            "./helpers/mastery.xlsm", // file path
            "Sheet1",                 // sheet tab name
            6,                        // starting row number
            "C",                      // marks column letter
            marks                     // marks array
        );
        res.status(200).send('PDF processed successfully.');
    } catch (err) {
        console.log(err);
        res.status(500).send('Error processing PDF.');
    }
}
module.exports = {
    uploadPdf
}