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

// Convert (a)(b)(c)(d) → 1/2/3/4
function letterToNumber(letter) {
    const map = { a: "1", b: "2", c: "3", d: "4" };
    return map[letter?.toLowerCase()] || "";
}

// Remove emojis only (do NOT remove line breaks)
function removeEmojis(text) {
    return text.replace(/[\p{Extended_Pictographic}]/gu, "");
}

// Clean spacing but preserve line breaks
function cleanText(text) {
    return text
        .replace(/\r/g, "")
        .replace(/[ \t]+/g, " ")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

// Convert multiline string to real DOCX paragraphs
function formatSolutionParagraphs(text) {
    text = text.replace(/\r/g, "").trim();
    text = text.replace(/(Statement\s*\d+\s*[-–]\s*)/gi, "\n$1");

    const parts = text.split("\n").filter(p => p.trim());

    return parts.map(part => new Paragraph(part.trim()));
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

            questionText = removeEmojis(questionText);
            questionText = cleanText(questionText);

            /* --------- OPTION EXTRACTION --------- */

            const options = [];

            const optionRegex = /\(([a-d])\)\s*([\s\S]*?)(?=\([a-d]\)|Answer|Correct\s*Answer|Explanation|$)/gi;

            let match;
            while ((match = optionRegex.exec(block)) !== null) {
                let optionText = match[2];
                optionText = removeEmojis(optionText);
                optionText = cleanText(optionText);
                options.push(optionText);
            }

            if (options.length > 4) options.splice(4);

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
                        new TableCell({ children: multilineParagraph(questionText) })
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
                            new TableCell({ children: multilineParagraph(opt) })
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
                        new TableCell({ children: formatSolutionParagraphs(explanationText) })
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