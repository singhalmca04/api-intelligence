const AdmZip = require("adm-zip");
const xml2js = require("xml2js");
const fs = require("fs");
const path = require("path");

async function updateMarksColumn(filePath, sheetName, startRow, columnLetter, marksArray, profileData) {
    try {
        const zip = new AdmZip(filePath);
        const parser = new xml2js.Parser();
        const builder = new xml2js.Builder();

        // ============================
        // STEP 1: Load workbook + rels
        // ============================
        const workbookXml = zip.readAsText("xl/workbook.xml");
        const workbookObj = await parser.parseStringPromise(workbookXml);

        const relsXml = zip.readAsText("xl/_rels/workbook.xml.rels");
        const relsObj = await parser.parseStringPromise(relsXml);

        // Find sheet mapping
        const sheetEntry = workbookObj.workbook.sheets[0].sheet
            .find(s => s.$.name.toLowerCase() === sheetName.toLowerCase());

        if (!sheetEntry) {
            throw new Error(`Sheet "${sheetName}" not found`);
        }

        const relId = sheetEntry.$["r:id"];
        const relation = relsObj.Relationships.Relationship
            .find(r => r.$.Id === relId);

        const sheetFile = `xl/${relation.$.Target}`;
        console.log("Updating sheet:", sheetFile);

        // ============================
        // STEP 2: Load sheet XML
        // ============================
        const sheetXml = zip.readAsText(sheetFile);
        const sheetObj = await parser.parseStringPromise(sheetXml);
        const sheetData = sheetObj.worksheet.sheetData[0];

        // ============================
        // STEP 3: Update marks column
        // ============================
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
                row.c.push({
                    $: { r: cellRef, t: "n" },
                    v: [mark.toString()]
                });
            } else {
                cell.$.t = "n";
                cell.v = [mark.toString()];
            }
        });

        const profileCellMap = {
            91: profileData.name || "",
            92: profileData.email || "",
            93: profileData.phone || "",
            94: profileData.dob || "",
            95: profileData.address || "",
            97: profileData.city || ""
        };

        Object.entries(profileCellMap).forEach(([rowNumber, value]) => {
            const rowRef = rowNumber.toString();
            const cellRef = `B${rowNumber}`;

            // Find or create row
            let row = sheetData.row.find(r => r.$.r === rowRef);
            if (!row) {
                row = { $: { r: rowRef }, c: [] };
                sheetData.row.push(row);
            }

            // Find or create cell
            let cell = row.c.find(c => c.$.r === cellRef);

            if (!cell) {
                row.c.push({
                    $: { r: cellRef, t: "str" },
                    v: [value]
                });
            } else {
                cell.$.t = "str";
                cell.v = [value];
            }
        });

        // ============================
        // STEP 4: Save updated sheet
        // ============================
        const updatedSheetXml = builder.buildObject(sheetObj);
        zip.updateFile(sheetFile, Buffer.from(updatedSheetXml, "utf8"));

        // =====================================================
        // STEP 5: FORCE EXCEL FORMULA RECALCULATION (CRITICAL)
        // =====================================================
        let updatedWorkbookXml = workbookXml;

        if (updatedWorkbookXml.includes("<calcPr")) {
            updatedWorkbookXml = updatedWorkbookXml.replace(
                /<calcPr[^>]*\/>/,
                '<calcPr fullCalcOnLoad="1"/>'
            );
        } else {
            updatedWorkbookXml = updatedWorkbookXml.replace(
                "</workbook>",
                '<calcPr fullCalcOnLoad="1"/></workbook>'
            );
        }

        zip.updateFile(
            "xl/workbook.xml",
            Buffer.from(updatedWorkbookXml, "utf8")
        );

        // ==================================
// STEP X: Update Sheet2 (C60–C66)
// ==================================
const sheet2Entry = workbookObj.workbook.sheets[0].sheet
    .find(s => s.$.name.toLowerCase() === "sheet2");

if (!sheet2Entry) {
    throw new Error("Sheet2 not found");
}

const sheet2RelId = sheet2Entry.$["r:id"];
const sheet2Relation = relsObj.Relationships.Relationship
    .find(r => r.$.Id === sheet2RelId);

const sheet2File = `xl/${sheet2Relation.$.Target}`;
console.log("Updating sheet:", sheet2File);

const sheet2Xml = zip.readAsText(sheet2File);
const sheet2Obj = await parser.parseStringPromise(sheet2Xml);
const sheet2Data = sheet2Obj.worksheet.sheetData[0];

// Write profile data to C60–C66
writeFixedCells(sheet2Data, {
    60: profileData.name || "",
    61: profileData.email || "",
    62: profileData.phone || "",
    63: profileData.dob || "",
    64: profileData.address || "",
    66: profileData.city || ""
}, "C");

// Save Sheet2
const updatedSheet2Xml = builder.buildObject(sheet2Obj);
zip.updateFile(sheet2File, Buffer.from(updatedSheet2Xml, "utf8"));

        // ============================
        // STEP 6: Save XLSM safely
        // ============================
        const outputFile = path.join(__dirname, "updated_mastery.xlsm");
        zip.writeZip(outputFile);

        console.log("✅ Marks updated and formulas will recalculate:", outputFile);

    } catch (err) {
        console.error("❌ Error:", err);
    }
}

// ---------------------------
// USAGE EXAMPLE
// ---------------------------
// const marks = [4, 4, 4, 4, 1, 2, 3]; // your marks array

// updateMarksColumn(
//     "./helpers/mastery.xlsm", // file path
//     "Sheet1",                 // sheet tab name
//     6,                        // starting row number
//     "C",                      // marks column letter
//     marks                     // marks array
// );

function extractProfileData(fileData) {
    const profile = {};

    fileData.forEach(line => {
        const parts = line.split(":");

        if (parts.length >= 2) {
            const key = parts[0].trim().toLowerCase();
            const value = parts.slice(1).join(":").trim();

            switch (key) {
                case "full name":
                    profile.name = value;
                    break;
                case "email":
                    profile.email = value;
                    break;
                case "phone number":
                    profile.phone = value;
                    break;
                case "dob":
                    profile.dob = value;
                    break;
                case "address":
                    profile.address = value;
                    break;
                case "city":
                    profile.city = value;
                    break;
            }
        }
    });

    return profile;
}
function writeFixedCells(sheetData, cellMap, columnLetter) {
    Object.entries(cellMap).forEach(([rowNumber, value]) => {
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
            row.c.push({
                $: { r: cellRef, t: "str" },
                v: [value]
            });
        } else {
            cell.$.t = "str";
            cell.v = [value];
        }
    });
}


module.exports = {
    updateMarksColumn,
    extractProfileData,
    writeFixedCells
};
