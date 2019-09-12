// Multiple cells merging in the same table
// Import from 'docx' rather than '../build' if you install from npm
import * as fs from "fs";
import { Document, Packer, Paragraph, Table, TableCell, TableRow } from "../build";

const doc = new Document();

const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("0,0")],
                }),
                new TableCell({
                    children: [new Paragraph("0,1")],
                    columnSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("0,3")],
                }),
                new TableCell({
                    children: [new Paragraph("0,4")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("1,0")],
                    columnSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("1,2")],
                    columnSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("1,4")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("2,0")],
                }),
                new TableCell({
                    children: [new Paragraph("2,1")],
                    columnSpan: 2,
                }),
                new TableCell({
                    children: [new Paragraph("2,3")],
                }),
                new TableCell({
                    children: [new Paragraph("2,4")],
                    columnSpan: 2,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("3,0")],
                }),
                new TableCell({
                    children: [new Paragraph("3,1")],
                }),
                new TableCell({
                    children: [new Paragraph("3,2")],
                }),
                new TableCell({
                    children: [new Paragraph("3,3")],
                }),
                new TableCell({
                    children: [new Paragraph("3,4")],
                }),
                new TableCell({
                    children: [new Paragraph("3,5")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph("4,0")],
                    columnSpan: 5,
                }),
                new TableCell({
                    children: [new Paragraph("4,5")],
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
                new TableCell({
                    children: [],
                }),
            ],
        }),
    ],
});

doc.addSection({
    children: [table],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
