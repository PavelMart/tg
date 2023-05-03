require("dotenv").config();
const fs = require("fs");
const TelegramBot = require("node-telegram-bot-api");
const { Configuration, OpenAIApi } = require("openai");
const Docx = require("./word/docx");

const TELEGRAM_BOT_API_KEY = process.env.TELEGRAM_BOT_API_KEY;

let apiKey = null;

const getResult = async (chatId, json) => {
  const configuration = new Configuration({
    apiKey,
  });

  const openai = new OpenAIApi(configuration);

  const sendCompletion = async (query) => {
    try {
      if (!query.handle) {
        return query.text;
      }
      const completion = await openai.createCompletion({
        model: "text-davinci-003",
        prompt: `${query.phrase}: ${query.text}`,
        temperature: 0.7,
        max_tokens: 2048,
      });
      return completion.data.choices[0].text.trim();
    } catch (error) {
      throw error;
    }
  };

  const handleParagraphs = async (paragraphs) => {
    const textList = paragraphs.map((p) => {
      const query = {
        handle: false,
        phrase: "",
        text: "",
      };
      if (!p["w:r"]) return query;

      const row = p["w:r"][0];

      const text = Array.isArray(row["w:t"]) ? row["w:t"].join("") : row["w:t"];

      if (p["w:pPr"][0].hasOwnProperty("w:jc") && p["w:pPr"][0]["w:jc"][0]["$"]["w:val"] === "center") return { ...query, text };

      if (row["w:rPr"][0].hasOwnProperty("w:highlight") && row["w:rPr"][0]["w:highlight"][0]["$"]["w:val"] === "green") {
        return { ...query, text };
      }
      return { handle: true, phrase: "Перефразируй", text };
    });

    const promisesList = textList.map(sendCompletion);

    const answersList = await Promise.all(promisesList);

    const result = paragraphs.map((p, i) => {
      if (answersList[i]) {
        const wt = answersList[i];

        const wr = [{ ...p["w:r"][0], ["w:t"]: [wt] }];

        return { ...p, "w:r": wr };
      }
      return p;
    });

    return result;
  };

  const { p: paragraphs } = Docx.getDocumentElements(json);

  const output = [];

  const count = 60;

  const pages = Math.floor(paragraphs.length / count);

  const delay = (ms) => new Promise((res) => setTimeout(res, ms));

  await bot.sendMessage(chatId, `Обработано 0/${paragraphs.length} параграфов`);

  for (let page = 0; page <= pages; page++) {
    let start = Date.now();
    const result = await handleParagraphs(paragraphs.slice(page * count, (page + 1) * count));
    output.push(...result);
    if (page < pages) {
      await bot.sendMessage(chatId, `Обработано ${(page + 1) * count}/${paragraphs.length} параграфов`);
      await delay(60000 - (Date.now() - start));
    } else await bot.sendMessage(chatId, `Обработано ${paragraphs.length}/${paragraphs.length} параграфов`);
  }

  return Docx.changeParagraphs(json, output);
};

const getReadStreamPromise = (stream, fileWriter) => {
  return new Promise((resolve, reject) => {
    stream.on("data", (chunk) => {
      fileWriter.write(chunk);
    });
    stream.on("error", (err) => {
      reject(err);
    });
    stream.on("end", () => {
      resolve();
    });
  });
};

const bot = new TelegramBot(TELEGRAM_BOT_API_KEY, { polling: true });

bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  try {
    if (msg.text && msg.text.includes("API_KEY=")) {
      apiKey = msg.text.split("=")[1];
      fs.writeFile("api.txt", apiKey, (err) => {
        if (err) throw err;
      });
    } else apiKey = fs.readFileSync("api.txt").toString();

    if (!apiKey) return await bot.sendMessage(chatId, "Пришлите API_KEY в формате: API_KEY=ваш_ключ_api");

    if (!msg.document) return await bot.sendMessage(chatId, "Пожалуйста, отправьте документ Word");

    const ext = msg.document.file_name.split(".")[1];

    if (ext !== "docx") return await bot.sendMessage(chatId, "Некорректный файл. Расширение обрабатываемого файла должно быть .docx");

    await bot.sendMessage(chatId, "Подождите, идет обработка файла");

    const fileWriter = fs.createWriteStream("file.docx");

    const stream = bot.getFileStream(msg.document.file_id);

    await getReadStreamPromise(stream, fileWriter);

    await bot.sendMessage(chatId, "Файл подготовлен, Отправляем данные в ChatGPT");

    Docx.extract("file.docx", "extracted");

    const originalJSON = Docx.translateDocumentXMLToJSON();

    const preparedJSON = Docx.prepareJSONForRephpasing(originalJSON);

    const resultJSON = await getResult(chatId, preparedJSON);

    Docx.translateDocumentJSONToXML(resultJSON);

    Docx.create("extracted", "result.docx");

    await bot.sendMessage(chatId, "Результат готов");

    const buffer = fs.readFileSync("result.docx");

    await bot.sendDocument(chatId, buffer, {}, { filename: "output.docx", contentType: "text" });
  } catch (error) {
    console.log(error);
    if (error.response && error.response.data.error.type === "insufficient_quota")
      bot.sendMessage(
        chatId,
        "Вы использовали всю доступную квоту, ChatGPT недоступен с данного аккаунта, смените API_KEY, введя: API_KEY=ваш_ключ_api"
      );
    else bot.sendMessage(chatId, "Произошла внутренняя ошибка, повторите позднее" + error);
  }
});

console.log("Bot is launched");
