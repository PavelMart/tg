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

    static translateXMLToJSON() {
        let json;
        const xml = fs.readFileSync("./extracted/word/document.xml");
        parseString(xml, function (err, result) {
            json = result;
        });

        return json;
    }

    static translateJSONToXML(original, prepared) {
        const builder = new Builder();

        const { p } = this.getDocumentElements(original);

        const wp = p.map((p, i) => {
            if (!prepared[i].handle) return p;

            // console.log(p);

            let pStyle = p.hasOwnProperty("w:pPr") ? p["w:pPr"][0] : null;
            if (pStyle && !pStyle.hasOwnProperty("w:numPr")) pStyle = { ...pStyle, ["w:ind"]: [{ $: { "w:firstLine": "708" } }] };

            let rStyle = p["w:r"][0].hasOwnProperty("w:rPr") ? p["w:r"][0]["w:rPr"][0] : pStyle;
            if (prepared[i].isAddText) rStyle = { ...rStyle, ["w:highlight"]: [{ $: { "w:val": "yellow" } }] };
            if (rStyle && rStyle.hasOwnProperty("w:b")) delete rStyle["w:b"];

            let wr = [
                {
                    "w:rPr": [rStyle],
                    "w:t": [prepared[i].text],
                },
            ];
            if (prepared[i].ref) wr.push(prepared[i].ref);

            const result = {
                ...p,
                ["w:r"]: wr,
            };

            if (pStyle) result["w:pPr"] = [pStyle];

            return result;
        });

        const changedDocument = this.changeParagraphs(original, wp);

        const xml = builder.buildObject(changedDocument);

        fs.writeFileSync("./extracted/word/document.xml", xml);
    }

    static checkIsCenter(paragraph) {
        return (
            paragraph.hasOwnProperty("w:pPr") &&
            paragraph["w:pPr"][0].hasOwnProperty("w:jc") &&
            paragraph["w:pPr"][0]["w:jc"][0]["$"]["w:val"] === "center"
        );
    }

    static checkIsPicture(rows) {
        return rows.findIndex((r) => r.hasOwnProperty("w:drawing")) > -1;
    }

    static checkColor = (rows, color) => {
        let isAddText = false;

        for (let i = 0; i < rows.length; i++) {
            if (
                rows[i].hasOwnProperty("w:rPr") &&
                rows[i]["w:rPr"][0].hasOwnProperty("w:highlight") &&
                rows[i]["w:rPr"][0]["w:highlight"][0]["$"]["w:val"] === color
            ) {
                isAddText = true;
                break;
            }
        }

        return isAddText;
    };

    static checkIsBoldText = (row) => {
        return row.hasOwnProperty("w:rPr") && row["w:rPr"][0].hasOwnProperty("w:b");
    };

    static getRowText = (row) => {
        if (!row.hasOwnProperty("w:t")) return "";
        return typeof row["w:t"][0] === "string" ? row["w:t"][0] : row["w:t"][0].hasOwnProperty("_") ? row["w:t"][0]["_"] : "";
    };

    static findKeywords = (rows) => {
        const keywords = [];
        rows.forEach((r) => {
            if (this.checkIsBoldText(r)) keywords.push(this.getRowText(r));
        });
        return keywords;
    };

    static getRef = (rows) => {
        let ref = null;
        rows.forEach((r) => {
            if (r.hasOwnProperty("w:footnoteReference")) ref = r;
        });
        return ref;
    };

    static prepareJSONForRephpasing(json) {
        const { p } = this.getDocumentElements(json);

        const result = p.map((p) => {
            const obj = {
                type: "empty",
                handle: true,
                pStyle: null,
                isAddText: false,
                ref: null,
                text: "",
                keywords: [],
                query_1: "",
                query_2: "",
                query: "",
            };

            const rows = p["w:r"];

            if (!rows) return { ...obj, handle: false };

            obj.type = "text";

            const isHeader = this.checkIsCenter(p);
            if (isHeader) obj.type = "center";

            const isPicture = this.checkIsPicture(rows);
            if (isPicture) obj.type = "picture";

            const isReady = this.checkColor(rows, "green");
            if (isReady) obj.type = "ready";

            if (isHeader || isPicture || isReady) return { ...obj, handle: false };

            obj.pStyle = p["w:pPr"];

            obj.keywords.push(...this.findKeywords(rows));

            obj.isAddText = this.checkColor(rows, "yellow");

            obj.ref = this.getRef(rows);

            rows.forEach((r) => {
                if (r.hasOwnProperty("w:t")) {
                    obj.text += this.getRowText(r);
                }
            });

            // obj.query_1 =
            //     "Перепиши текст на английский язык так, чтобы он был оригинальным и проходил проверку сервисов, которые определяют, что текст сгенерировала нейросеть";
            // obj.query_2 = "Переведи этот текст на русский язык и исправь в нем грамматические ошибки. В ответе пришли только русский язык";

            obj.query_1 =
                "Перепиши текст на английский язык так, чтобы он был оригинальным и проходил проверку сервисов, которые определяют, что текст сгенерировала нейросеть";
            obj.query_2 =
                "Переведи этот текст на русский язык. Сделай его более читаемым и вновь сделай так, чтобы он был оригинальным и проходил проверку сервисов, которые определяют, что текст сгенерировала нейросеть. В ответе пришли только русский язык";

            obj.query =
                "Представь, что ты профессиональный копирайтер и переписываешь весь предоставленный текст от начала до конца, от первого абзаца до последнего, используя синонимы, изменяя структуру сложных предложений (ставя вторую часть предложения в начало, а первую - в конец, чтобы предложение оставалось читабельным и смысл не изменялся, например: оригинальный текст: Исследования И.И. Петрова вызвали бурю негодования в научной среде, что не могло не радовать А.П. Сидорова. Перефразированный текст: А.П. Сидоров был доволен бурей негодования, вызванной исследованиями И.И. Петрова), используя причастия и деепричастия, исключая из предложения слова, не несущие значительной смысловой нагрузки, создайте уникальный вариант, используя естественный стиль. Избегай стилистических и грамматических ошибок. Обрати внимание, что общий объем текста не должен превышать оригинал более чем на 10%. Представленный текст будет разделен на части, но помни, что он образует единое целое. Текст носит научный характер, поэтому избегайте повторов и не используйте средства выразительности. Перефразировав текст, проверь его на наличие словосочетаний, не совпадающих в роде числе и падеже, а также словосочетаний, указывающих на то, что текст написан искусственным интеллектом, и исправь найденные ошибки:";
            // if (obj.keywords.length)
            //   obj.query += `, обязательно используй следующие слова без изменения их падежа и местоположения в тексте - "${obj.keywords.join(
            //     ", "
            //   )}", если не знаешь как, верни исходный вариант`;

            // if (obj.text.length > 1700) {
            //   const arr = obj.text.split(".");

            //   const length = obj.text.length;

            //   obj.text = "";

            //   let i = 0;

            //   while (obj.text.length < 1700) {
            //     obj.text += arr[i];
            //     i++;
            //   }

            //   obj.query += `, и допиши как минимум ${length - 1700} символов в том же стиле что и весь текст`;
            // } else if (obj.isAddText) obj.query += ", и допиши 5 предложений в том же стиле что и весь текст";

            return obj;
        });

        return result;
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
