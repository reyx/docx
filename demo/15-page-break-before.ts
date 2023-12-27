// Page break before example

import { Document, Packer, Paragraph } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph("Hello World"),
                new Paragraph({
                    text: "Hello World on another page",
                    pageBreakBefore: true,
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
