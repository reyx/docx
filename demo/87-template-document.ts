// Patch a document with patches

import { patchDocument, PatchType, TextRun } from "@reyx/docx";
import * as fs from "fs";

patchDocument(fs.readFileSync("demo/assets/simple-template-2.docx"), {
    patches: {
        name: {
            type: PatchType.PARAGRAPH,
            children: [new TextRun("Max")],
        },
    },
}).then((doc) => {
    fs.writeFileSync("My Document.docx", doc);
});
