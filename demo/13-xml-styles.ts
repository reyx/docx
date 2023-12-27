// This example shows 3 styles using XML styles

import { Document, HeadingLevel, Packer, Paragraph } from "@reyx/docx";
import * as fs from "fs";

const styles = fs.readFileSync("./demo/assets/custom-styles.xml", "utf-8");
const doc = new Document({
    title: "Title",
    externalStyles: styles,
    sections: [
        {
            children: [
                new Paragraph({
                    text: "Cool Heading Text",
                    heading: HeadingLevel.HEADING_1,
                }),
                new Paragraph({
                    text: 'This is a custom named style from the template "MyFancyStyle"',
                    style: "MyFancyStyle",
                }),
                new Paragraph("Some normal text"),
                new Paragraph({
                    text: "MyFancyStyle again",
                    style: "MyFancyStyle",
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
