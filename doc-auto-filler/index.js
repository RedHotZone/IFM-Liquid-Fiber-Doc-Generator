console.log("=== Node Automation Engine Initializing ===");

const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ExcelJS = require('exceljs');
const multer = require('multer');
const axios = require('axios');
const FormData = require('form-data');

const app = express();
const upload = multer();

app.use(express.urlencoded({ extended: true })); 
app.use(express.json());

// n8n production webhook URL
const N8N_WEBHOOK_URL = process.env.N8N_WEBHOOK_URL || "https://red-hot-zone.app.n8n.cloud/webhook/8f803c64-7a6f-4d04-9169-fe28f8367015";

// 🎯 Static Form Registry & Config Modules (Ensures clean Vercel bundling)
const FORM_CONFIGS = {
    "262012074627046": require('./Forms/manhole_access.js'),
    "262001795087054": require('./Forms/word_sample.js'),
    // Add future forms here: "FORM_ID": require('./Forms/your_file.js')
};

app.post('/generate-doc', upload.none(), async (req, res) => {
    // 1. Instantly respond to Jotform
    res.status(200).send("Webhook received, processing automation pipeline...");

    try {
        const { formID, pretty: prettyString = "" } = req.body || {};
        console.log(`📥 Processing Form ID: ${formID}`);

        // Look up configuration module
        const formConfig = FORM_CONFIGS[formID];

        if (!formConfig) {
            return console.error(`Error: Form ID ${formID} has no registered configuration module.`);
        }

        const { projectName, submissionDate, templateData } = formConfig.parseFields(prettyString);
        
        // Extract template base name without extension 
        const templateBaseName = path.parse(formConfig.templateFile).name;

        const dynamicFilename = `${projectName} - ${templateBaseName} - ${submissionDate}${formConfig.fileExtension}`;
        const folderName = formConfig.type === "word" ? "Word Templates" : "Excel Templates";
        const templatePath = path.resolve(__dirname, 'Templates', folderName, formConfig.templateFile);

        let fileBuffer;

        // --- BRANCH A: COMPILING WORD DOCUMENT ---
        if (formConfig.type === "word") {
            const content = fs.readFileSync(templatePath, 'binary');
            const doc = new Docxtemplater(new PizZip(content), { 
                paragraphLoop: true, 
                linebreaks: true, 
                delimiters: { start: "{{", end: "}}" } 
            });
            
            doc.render(templateData);
            fileBuffer = doc.toBuffer();
        }

        // --- BRANCH B: COMPILING EXCEL SPREADSHEET ---
        if (formConfig.type === "excel") {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(templatePath);

            Object.keys(templateData).forEach(sheetNameOrCell => {
                const val = templateData[sheetNameOrCell];

                if (typeof val === 'object' && val !== null) {
                    // Multi-sheet structure: { "Sheet1": { "A1": "Val" } }
                    const worksheet = workbook.getWorksheet(sheetNameOrCell);

                    if (!worksheet) {
                        console.warn(`Warning: Worksheet "${sheetNameOrCell}" not found in template. Skipping.`);
                        return;
                    }

                    console.log(`📑 Writing data to sheet: "${worksheet.name}"`);

                    Object.keys(val).forEach(cell => {
                        worksheet.getCell(cell).value = val[cell];
                    });
                } else {
                    // Single-sheet fallback structure: { "A1": "Val" }
                    const worksheet = workbook.worksheets[0];
                    worksheet.getCell(sheetNameOrCell).value = val;
                }
            });

            fileBuffer = await workbook.xlsx.writeBuffer();
        }

        // 🚀 2. Bundle Document + Metadata for n8n
        console.log(` Bundling ${dynamicFilename} and sending to n8n...`);
        const form = new FormData();
        
        form.append('file', fileBuffer, {
            filename: dynamicFilename,
            contentType: formConfig.contentType
        });

        form.append('projectName', projectName);
        form.append('submissionDate', submissionDate);
        form.append('formID', formID);

        // 3. Fire payload to n8n
        const response = await axios.post(N8N_WEBHOOK_URL, form, {
            headers: { ...form.getHeaders() }
        });

        console.log(`Success! n8n workflow invoked flawlessly. Response status: ${response.status}`);

    } catch (error) {
        console.error("Automation pipeline error:", error.message);
    }
});

// Dynamic Port Assignment (Supports Vercel & Local Dev)
const PORT = process.env.PORT || 3000;

if (process.env.NODE_ENV !== 'production') {
    app.listen(PORT, () => console.log(`Automation Engine running on port ${PORT}`));
}

module.exports = app;