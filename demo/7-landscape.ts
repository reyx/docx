// Example of how to set the document to landscape

import { Document, Packer, PageOrientation, Paragraph } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            properties: {
                page: {
                    size: {
                        orientation: PageOrientation.LANDSCAPE,
                    },
                },
            },
            children: [new Paragraph("Hello World")],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
