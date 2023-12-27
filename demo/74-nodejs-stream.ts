// Exporting the document as a stream

import { Document, Packer, Paragraph, Tab, TextRun } from "@reyx/docx";
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
                        new TextRun({
                            children: [new Tab(), "Github is the best"],
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
    ],
});

const stream = Packer.toStream(doc);
stream.pipe(fs.createWriteStream("My Document.docx"));
