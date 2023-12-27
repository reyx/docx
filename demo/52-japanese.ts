// Japanese text - Need to use a Japanese font

import { Document, HeadingLevel, Packer, Paragraph } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    styles: {
        paragraphStyles: [
            {
                id: "Normal",
                name: "Normal",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "MS Gothic",
                },
            },
        ],
    },
    sections: [
        {
            children: [
                new Paragraph({
                    text: "KFCを食べるのが好き",
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    text: "こんにちは",
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
