const path = require("path");
const AdmZip = require("adm-zip");
const xml2js = require("xml2js");
const fs = require("fs");

exports.updateXlsmSafe = async (req, res) => {
    try {
        // 1. Load original XLSM file
        const filePath = path.join(__dirname, "../uploads/mastery.xlsm");
        const zip = new AdmZip(filePath);

        // 2. Extract sheet1.xml (or the sheet you want)
        const sheetEntry = zip.getEntry("xl/worksheets/sheet1.xml");

        if (!sheetEntry) {
            return res.status(400).json({ error: "Sheet not found!" });
        }

        const sheetXml = sheetEntry.getData().toString("utf8");

        // 3. Parse XML
        const parser = new xml2js.Parser();
        const builder = new xml2js.Builder();

        const sheetObj = await parser.parseStringPromise(sheetXml);

        // -----------------------------------------------------
        // 4. ADD DATA SAFELY (modify rows without touching macros)
        // -----------------------------------------------------
        const users = [
            { name: "Amit", marks: 80, Score: "101" },
            { name: "Sumit", marks: 95, Score: "102" }
        ];

        const sheetData = sheetObj.worksheet.sheetData[0];

        // Find last row number
        let lastRow = sheetData.row ? sheetData.row.length : 0;
        console.log("lastRow " + lastRow);
        users.forEach((user) => {
            lastRow++;

            sheetData.row.push({
                $: { r: lastRow },
                c: [
                    { _: lastRow, $: { r: `A${lastRow}`, t: "n" } },
                    { _: user.name, $: { r: `B${lastRow}`, t: "str" } },
                    { _: user.marks, $: { r: `C${lastRow}`, t: "n" } },
                    { _: user.Score, $: { r: `D${lastRow}`, t: "str" } }
                ]
            });
        });

        // 5. Rebuild updated XML
        const updatedSheetXml = builder.buildObject(sheetObj);

        // 6. Replace ONLY sheet1.xml (macros remain untouched)
        zip.updateFile("xl/worksheets/sheet1.xml", Buffer.from(updatedSheetXml, "utf8"));

        // 7. Save as a NEW XLSM file
        const outputPath = path.join(__dirname, "../uploads/students_updated.xlsm");
        zip.writeZip(outputPath);

        res.json({
            message: "Macro-safe XLSM updated successfully!",
            file: "students_updated.xlsm"
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Failed to update XLSM safely" });
    }
};