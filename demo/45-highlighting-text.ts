// Highlighting text

import { AlignmentType, Document, Header, Packer, Paragraph, TextRun } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    text: "Hello World",
                                    color: "FF0000",
                                    bold: true,
                                    size: 24,
                                    font: {
                                        name: "Garamond",
                                    },
                                    highlight: "yellow",
                                }),
                            ],
                        }),
                    ],
                }),
            },
            children: [],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
