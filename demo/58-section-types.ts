// Usage of different Section Types

import { Document, Packer, Paragraph, SectionType, TextRun } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
        {
            properties: {
                type: SectionType.CONTINUOUS,
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
        {
            properties: {
                type: SectionType.ODD_PAGE,
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
        {
            properties: {
                type: SectionType.EVEN_PAGE,
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
        {
            properties: {
                type: SectionType.NEXT_PAGE,
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
