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

// Convert letter → number
function letterToNumber(letter) {
    const map = { a: "1", b: "2", c: "3", d: "4" };
    return map[letter?.toLowerCase()] || "";
}

app.post("/upload-doc", upload.single("file"), async (req, res) => {
    try {
        const filePath = req.file.path;

        const result = await mammoth.extractRawText({ path: filePath });
        const text = result.value;

        // Split by Q1. Q2. Q3.
        const blocks = text.split(/\n?\s*Q\s*\d+[\.\)\:\-]?\s*/i).filter(b => b.trim());

        const children = [];

        blocks.forEach(block => {

            // Extract Question
            const questionPart = block.split("(a)")[0].trim();
            const questionText = questionPart.trim();

            // Extract Options
            const options = [];
            const optionRegex = /\([a-d]\)\s*(.+)/gi;
            let match;
            while ((match = optionRegex.exec(block)) !== null) {
                options.push(match[1].trim());
            }

            // Extract Answer
            const answerMatch = block.match(/Answer\s*[:\-]?\s*\(?([a-d])\)?/i);
            const answerLetter = answerMatch ? answerMatch[1] : "";
            const answerNumber = letterToNumber(answerLetter);

            // Extract Explanation
            let explanationText = "";
            const explanationSplit = block.split(/Explanation/i);
            if (explanationSplit.length > 1) {
                explanationText = explanationSplit[1].trim();
            }

            // Build Table Rows
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

            // Solution (Blank)
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
                rows: rows,
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

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log("Server running on port " + PORT);
});