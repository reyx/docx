// Add text to header and footer

import { Document, Footer, Header, Packer, Paragraph } from "@reyx/docx";
import * as fs from "fs";

const doc = new Document({
    sections: [
        {
            headers: {
                default: new Header({
                    children: [new Paragraph("Header text")],
                }),
            },
            footers: {
                default: new Footer({
                    children: [new Paragraph("Footer text")],
                }),
            },
            children: [new Paragraph("Hello World")],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
