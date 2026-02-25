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

// Convert option letter → number
function letterToNumber(letter) {
    const map = { a: "1", b: "2", c: "3", d: "4" };
    return map[letter?.toLowerCase()] || "";
}

// Remove emojis
function removeEmojis(text) {
    return text.replace(/[\p{Extended_Pictographic}]/gu, "");
}

// Clean extra spacing
function cleanText(text) {
    return text.replace(/\s+/g, " ").trim();
}

/* ------------------------- MAIN ROUTE ------------------------- */

app.post("/upload-doc", upload.single("file"), async (req, res) => {
    try {
        const filePath = req.file.path;

        const result = await mammoth.extractRawText({ path: filePath });
        const text = result.value;

        // Robust split: handles Q1., Q1:, Q1), Q 1 etc.
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

            /* --------- OPTION EXTRACTION --------- */

            const options = [];

            // Find all option markers like (a), (b), (c), (d)
            const optionMatches = [...block.matchAll(/\([a-d]\)/gi)];

            for (let i = 0; i < optionMatches.length; i++) {
                const start = optionMatches[i].index;
                const end = (i + 1 < optionMatches.length)
                    ? optionMatches[i + 1].index
                    : block.search(/(Answer|Correct Answer|Explanation)/i) > -1
                        ? block.search(/(Answer|Correct Answer|Explanation)/i)
                        : block.length;

                const optionText = block
                    .substring(start)
                    .replace(/\([a-d]\)/i, "")
                    .trim();

                options.push(optionText.substring(0, end - start).trim());
            }

            /* --------- ANSWER EXTRACTION --------- */

            const answerMatch = block.match(/(Correct\s*)?Answer\s*[:\-]?\s*\(?([a-d])\)?/i);

            const answerLetter = answerMatch ? answerMatch[2] : "";
            const answerNumber = letterToNumber(answerLetter);

            /* --------- EXPLANATION EXTRACTION --------- */

            let explanationText = "";

            const explanationMatch = block.split(/Explanation/i);

            if (explanationMatch.length > 1) {
                explanationText = explanationMatch[1];
                explanationText = removeEmojis(explanationText);
                explanationText = cleanText(explanationText);
            }

            /* --------- BUILD TABLE --------- */

            const rows = [];

            // Question
            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Question")] }),
                        new TableCell({ children: [new Paragraph(questionText)] })
                    ]
                })
            );

            // Type
            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Type")] }),
                        new TableCell({ children: [new Paragraph("multiple_choice")] })
                    ]
                })
            );

            // Options
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

            // Answer
            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Answer")] }),
                        new TableCell({ children: [new Paragraph(answerNumber)] })
                    ]
                })
            );

            // Solution (blank)
            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Solution")] }),
                        new TableCell({ children: [new Paragraph(explanationText)] })
                    ]
                })
            );


            // Positive Marks
            rows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Positive Marks")] }),
                        new TableCell({ children: [new Paragraph("2")] })
                    ]
                })
            );

            // Negative Marks
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