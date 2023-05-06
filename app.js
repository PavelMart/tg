const Docx = require("./word/docx");
const fs = require("fs");

Docx.extract("file.docx", "extracted");

const originalJSON = Docx.translateXMLToJSON();

fs.writeFileSync("o.json", JSON.stringify(originalJSON));

const preparedJSON = Docx.prepareJSONForRephpasing(originalJSON);

fs.writeFileSync("p.json", JSON.stringify(preparedJSON));

Docx.translateJSONToXML(originalJSON, preparedJSON);

// const json = fs.readFileSync("o.json").toString();

// Docx.translateDocumentJSONToXML(JSON.parse(json));

Docx.create("extracted", "result.docx");
