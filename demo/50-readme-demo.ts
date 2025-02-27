// The demo on the README.md

import { Document, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, VerticalAlign } from "@reyx/docx";
import * as fs from "fs";

const table = new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./demo/images/image1.jpeg"),
                                    transformation: {
                                        width: 100,
                                        height: 100,
                                    },
                                }),
                            ],
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            text: "Hello",
                            heading: HeadingLevel.HEADING_1,
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            text: "World",
                            heading: HeadingLevel.HEADING_1,
                        }),
                    ],
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./demo/images/image1.jpeg"),
                                    transformation: {
                                        width: 100,
                                        height: 100,
                                    },
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        }),
    ],
});

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    text: "Hello World",
                    heading: HeadingLevel.HEADING_1,
                }),
                table,
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./demo/images/pizza.gif"),
                            transformation: {
                                width: 100,
                                height: 100,
                            },
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
