const AdmZip = require("adm-zip");
const fs = require("fs");
const { parseString, Builder } = require("xml2js");

class Docx {
  static extract(from, to) {
    const zip = new AdmZip(from);
    zip.extractAllTo(to, true);
  }

  static create(from, to) {
    const zip = new AdmZip();

    zip.addLocalFolder(from);

    zip.writeZip(to);

    fs.rm("./extracted", { recursive: true }, (err) => {
      if (err) console.error(err);
    });
  }

  static translateDocumentXMLToJSON() {
    let json;
    const xml = fs.readFileSync("./extracted/word/document.xml");
    parseString(xml, function (err, result) {
      json = result;
    });

    return json;
  }

  static translateDocumentJSONToXML(json) {
    const builder = new Builder();

    const xml = builder.buildObject(json);

    fs.writeFileSync("./extracted/word/document.xml", xml);
  }

  static prepareJSONForRephpasing(json) {
    const { p } = this.getDocumentElements(json);

    const wp = p.map((p) => {
      if (!p["w:r"]) return p;
      if (!p["w:pPr"][0]["w:rPr"]) return p;

      if (p["w:pPr"][0]["w:rPr"][0].hasOwnProperty("w:b")) {
        return p;
      }

      const wt = p["w:r"]
        .map((r) => {
          if (r["w:t"]) return r["w:t"][0].hasOwnProperty("_") ? r["w:t"][0]["_"] : r["w:t"][0];
        })
        .filter((r) => typeof r === "string")
        .join("");

      const wr = [{ ...p["w:r"][0], ["w:t"]: wt }];

      const wp = { ...p, "w:r": wr };

      return wp;
    });

    return this.changeParagraphs(json, wp);
  }

  static changeParagraphs(json, wp) {
    const { document, body } = this.getDocumentElements(json);

    return { ...json, ["w:document"]: { ...document, ["w:body"]: [{ ...body, ["w:p"]: wp }] } };
  }

  static getDocumentElements(json) {
    const document = json["w:document"];

    const body = document["w:body"][0];

    const p = body["w:p"];

    return { document, body, p };
  }
}

module.exports = Docx;
