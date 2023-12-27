// Section with 2 columns including a column break

import { ColumnBreak, Document, Packer, Paragraph, TextRun } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            properties: {
                column: {
                    space: 708,
                    count: 2,
                },
            },
            children: [
                new Paragraph({
                    children: [
                        new TextRun("This text will be in the first column."),
                        new ColumnBreak(),
                        new TextRun("This text will be in the second column."),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
