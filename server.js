const express = require("express");

const fs = require("fs");
const cors = require("cors");

const path = require("path");

const PizZip = require("pizzip");

const Docxtemplater = require("docxtemplater");

const app = express();

app.use(express.json());

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "https://dhr-helper.netlify.app");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  next();
});

app.use(cors({
  origin: ["http://localhost:5173","https://dhr-helper.netlify.app"],
  methods: ["GET", "POST", "OPTIONS"],
  allowedHeaders: ["Content-Type"]
}));

app.options("*",cors());
const filePath = path.join(__dirname, "template", "release-template.docx");

const TEMPLATE_PATH = filePath;

function extractRequestFields(obj) {
  let rows = [];

  function traverse(current, parent = "") {
    for (let key in current) {
      const value = current[key];

      const fieldName = parent ? `${parent}.${key}` : key;

      let type = typeof value;

      if (Array.isArray(value)) {
        type = "array";
      } else if (type === "object" && value !== null) {
        type = "object";
      }

      rows.push({
        sno: rows.length + 1,

        field: fieldName,

        type: type,

        mandatory: "M",

        description: "",
      });

      if (type === "object") {
        traverse(value, fieldName);
      }
    }
  }

  traverse(obj);

  return rows;
}

// -------- API ----------

app.post("/generate-release-notes", (req, res) => {
  const {
    version,

    date,

    apiName,

    url,

    request,

    response,

    environment,

    remarks,

    Description,
  } = req.body;

  try {
    const content = fs.readFileSync(TEMPLATE_PATH, "binary");

    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,

      linebreaks: true,
    });

    // 🔹 Convert request JSON → professional table

    const requestTable = extractRequestFields(request);

    doc.setData({
      version,

      date,

      apiName,

      Description,

      url,

      request: JSON.stringify(request, null, 2),

      response: JSON.stringify(response, null, 2),

      environment,

      remarks,

      requestTable,
    });

    doc.render();

    const buf = doc.getZip().generate({
      type: "nodebuffer",

      compression: "DEFLATE",
    });


    const fileName = `${apiName}-release-notes-${version}.docx`;

    // Send the file directly without saving to disk
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.send(buf);
  } catch (error) {
    console.error(error);

    res.status(500).json({
      message: "Error generating release notes",

      error: error.message,
    });
  }
});

module.exports =app;


// app.listen(5000, () => {
//   console.log("Release Notes Generator running on port 5000");
// });
