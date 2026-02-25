const express = require("express");
const cors = require("cors");
const multer = require("multer");
const mammoth = require("mammoth");
const path = require("path");

const {
    Document,
    Packer,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    WidthType
} = require("docx");

const app = express();

app.use(cors());
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ dest: "uploads/" });

/* ------------------------- UTILITIES ------------------------- */

function letterToNumber(letter) {
    const map = { a: "1", b: "2", c: "3", d: "4" };
    return map[letter?.toLowerCase()] || "";
}

function removeEmojis(text) {
    return text.replace(/[\p{Extended_Pictographic}]/gu, "");
}

function cleanText(text) {
    return text.replace(/\s+/g, " ").trim();
}

/* ------------------------- MAIN ROUTE ------------------------- */

app.post("/upload-doc", upload.single("file"), async (req, res) => {
    try {
        const filePath = req.file.path;

        const result = await mammoth.extractRawText({ path: filePath });
        const text = result.value;

        const blocks = text
            .split(/\n?\s*(?=Q\s*\d+)/i)
            .filter(b => b.trim());

        console.log("Detected questions:", blocks.length);

        const children = [];

        blocks.forEach(block => {

            /* --------- QUESTION EXTRACTION --------- */

            const firstOptionMatch = block.match(/\([a-d]\)/i);

            let questionText = "";

            if (firstOptionMatch) {
                questionText = block
                    .substring(0, firstOptionMatch.index)
                    .replace(/^Q\s*\d+[\.\:\-\)]?\s*/i, "");
            } else {
                questionText = block;
            }

            questionText = cleanText(questionText);

            /* --------- OPTION EXTRACTION (FIXED) --------- */

            const options = [];

            const optionRegex = /\(([a-d])\)\s*([\s\S]*?)(?=\([a-d]\)|Answer|Correct\s*Answer|Explanation|$)/gi;

            let match;

            while ((match = optionRegex.exec(block)) !== null) {
                let optionText = match[2];

                optionText = removeEmojis(optionText);
                optionText = cleanText(optionText);

                options.push(optionText);
            }

            // Strict 4 options safeguard
            if (options.length > 4) {
                options.splice(4);
            }

            /* --------- ANSWER EXTRACTION --------- */

            const answerMatch = block.match(
                /(Correct\s*)?Answer\s*[:\-]?\s*\(?([a-d])\)?/i
            );

            const answerLetter = answerMatch ? answerMatch[2] : "";
            const answerNumber = letterToNumber(answerLetter);

            /* --------- EXPLANATION EXTRACTION --------- */

            let explanationText = "";

            const explanationRegex = /Explanation\s*[:\-]?\s*([\s\S]*)/i;
            const explanationMatch = block.match(explanationRegex);

            if (explanationMatch) {
                explanationText = removeEmojis(explanationMatch[1]);
                explanationText = cleanText(explanationText);
            }

            /* --------- BUILD TABLE --------- */

            const rows = [];

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Question")] }),
                        new TableCell({ children: [new Paragraph(questionText)] })
                    ]
                })
            );

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Type")] }),
                        new TableCell({ children: [new Paragraph("multiple_choice")] })
                    ]
                })
            );

            options.forEach(opt => {
                rows.push(
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("Option")] }),
                            new TableCell({ children: [new Paragraph(opt)] })
                        ]
                    })
                );
            });

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Answer")] }),
                        new TableCell({ children: [new Paragraph(answerNumber)] })
                    ]
                })
            );

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Solution")] }),
                        new TableCell({ children: [new Paragraph(explanationText)] })
                    ]
                })
            );

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Positive Marks")] }),
                        new TableCell({ children: [new Paragraph("2")] })
                    ]
                })
            );

            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Negative Marks")] }),
                        new TableCell({ children: [new Paragraph("0.66")] })
                    ]
                })
            );

            const table = new Table({
                rows,
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                }
            });

            children.push(table);
            children.push(new Paragraph(""));
        });

        /* --------- GENERATE DOC --------- */

        const doc = new Document({
            sections: [{ children }]
        });

        const buffer = await Packer.toBuffer(doc);

        res.setHeader(
            "Content-Disposition",
            "attachment; filename=Converted_Output.docx"
        );

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );

        res.send(buffer);

    } catch (err) {
        console.error(err);
        res.status(500).send("Error processing file");
    }
});

/* ------------------------- SERVER ------------------------- */

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
    console.log("Server running on port " + PORT);
});