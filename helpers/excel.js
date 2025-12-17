const AdmZip = require("adm-zip");
const xml2js = require("xml2js");
const fs = require("fs");
const path = require("path");

async function updateMarksColumn(filePath, sheetName, startRow, columnLetter, marksArray) {
    try {
        const zip = new AdmZip(filePath);

        // STEP 1: Load workbook + relationships
        const workbookXml = zip.readAsText("xl/workbook.xml");
        const workbookObj = await new xml2js.Parser().parseStringPromise(workbookXml);

        const relsXml = zip.readAsText("xl/_rels/workbook.xml.rels");
        const relsObj = await new xml2js.Parser().parseStringPromise(relsXml);

        // Find sheet mapping (Sheet1 → sheet1.xml)
        const sheetEntry = workbookObj.workbook.sheets[0].sheet
            .find(s => s.$.name.toLowerCase() === sheetName.toLowerCase());

        const relId = sheetEntry.$["r:id"];
        const relation = relsObj.Relationships.Relationship.find(r => r.$.Id === relId);

        const sheetFile = `xl/${relation.$.Target}`;
        console.log("Updating sheet:", sheetFile);

        // STEP 2: Load sheet XML
        const sheetXml = zip.readAsText(sheetFile);
        const sheetObj = await new xml2js.Parser().parseStringPromise(sheetXml);

        const sheetData = sheetObj.worksheet.sheetData[0];

        // STEP 3: Loop through marks array
        marksArray.forEach((mark, index) => {
            const rowNumber = startRow + index;
            const rowRef = rowNumber.toString();
            const cellRef = `${columnLetter}${rowNumber}`;

            // Find or create row
            let row = sheetData.row.find(r => r.$.r === rowRef);
            if (!row) {
                row = { $: { r: rowRef }, c: [] };
                sheetData.row.push(row);
            }

            // Find or create cell
            let cell = row.c.find(c => c.$.r === cellRef);

            if (!cell) {
                // Create new numeric cell
                row.c.push({
                    _: mark,
                    $: { r: cellRef, t: "n" },
                    v: [mark.toString()]
                });
            } else {
                // Force cell to numeric type
                cell.$.t = "n";

                if (cell.v) {
                    cell.v[0] = mark.toString();
                } else {
                    cell.v = [mark.toString()];
                }
            }
        });

        // STEP 4: Rebuild XML
        const updatedXml = new xml2js.Builder().buildObject(sheetObj);
        zip.updateFile(sheetFile, Buffer.from(updatedXml));

        // STEP 5: Save output XLSM
        const outputFile = path.join(__dirname, "updated_mastery.xlsm");
        zip.writeZip(outputFile);

        console.log("Marks updated successfully:", outputFile);

    } catch (err) {
        console.error("Error:", err);
    }
}

// ---------------------------
// USAGE EXAMPLE
// ---------------------------
// const marks = [1, 2, 3, 4, 1, 2, 3]; // your marks array

// updateMarksColumn(
//     "./helpers/mastery.xlsm", // file path
//     "Sheet1",                 // sheet tab name
//     6,                        // starting row number
//     "C",                      // marks column letter
//     marks                     // marks array
// );

module.exports = {
    updateMarksColumn
}