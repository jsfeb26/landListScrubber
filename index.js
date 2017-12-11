const commandLineArgs = require("command-line-args");
const XLSX = require("xlsx");
const fetch = require("node-fetch");
const get = require("lodash.get");
const _progress = require("cli-progress");

const NEW_DESCRIPTION_FIELDS = [
  { name: "Section", match: ["section:"] },
  { name: "Township", match: ["township:"] },
  { name: "Range", match: ["range:"] },
  { name: "Acres", match: ["acres", "acre"] }
];
const ASSESSOR_FIELDS = [
  { name: "Tax Year", match: "taxYear" },
  { name: "GIS Map URL", match: "gisMapURL" },
  { name: "Land Value", match: "landValue" },
  { name: "Full Cash Value", match: "fullCashValue" },
  { name: "Assessed Full Cash Value", match: "assessedFullCashValue" },
  { name: "Assessed Limited Value", match: "assessedLimitedValue" },
  { name: "Sale Price", match: "salePrice" },
  { name: "Sale Date", match: "saleDate" },
  { name: "Assessed Full Cash Value Amount", match: "assessedFullCashValueAmount" },
  { name: "Assessed Limited Value Amount", match: "assessedLimitedValueAmount" },
  { name: "Multiple Owners", match: "multipleOwners" }
];
const ASSESSOR_URL = "https://www.mohavecounty.us/service/assessor_parcel/search/numbers/";

const optionDefinitions = [{ name: "file", type: String, multiple: false, defaultOption: true }];

const fetchAssessorData = async parcel => {
  const res = await fetch(
    `${ASSESSOR_URL}${parcel.slice(0, 3)}-${parcel.slice(3, 5)}-${parcel.slice(5)}`
  );
  const data = await res.json();
  const entries = get(data, "data.entrys") || [];
  return entries[entries.length - 1] || {};
};

const processWorksheet = async worksheet => {
  const newWorksheet = worksheet.map(
    (property, index) =>
      new Promise((resolve, reject) => {
        setTimeout(async () => {
          const descriptionParts = (property["Legal Desc"] || "").split(" ");

          NEW_DESCRIPTION_FIELDS.forEach(field => {
            const partIndex = descriptionParts.findIndex(
              part => field.match.indexOf(part.toLowerCase()) > -1
            );

            let fieldVal = "";
            if (partIndex > -1) {
              if (field.name === "Acres") {
                fieldVal = descriptionParts[partIndex - 1];
              } else {
                fieldVal = descriptionParts[partIndex + 1];
              }
            }

            property[field.name] = fieldVal;
          });

          const assessorData = await fetchAssessorData(property["Parcel."]);
          ASSESSOR_FIELDS.forEach(field => {
            property[field.name] = assessorData[field.match] || "";
          });

          resolve(property);
        }, index * 1500);
      })
  );

  return await Promise.all(newWorksheet);
};

(async function() {
  const options = commandLineArgs(optionDefinitions);
  const workbook = XLSX.readFile(options.file);
  const newWorkbook = XLSX.readFile("new_mohave.xlsx");

  const sheetName = workbook.SheetNames[0];
  const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  let newWorksheet = XLSX.utils.sheet_to_json(newWorkbook.Sheets[sheetName]) || [];

  let prevIndex = newWorksheet.length;
  let currIndex = prevIndex + 100;
  const total = worksheet.length;

  const progressBar = new _progress.Bar({}, _progress.Presets.shades_classic);
  progressBar.start(total, prevIndex);

  do {
    newWorksheet = [
      ...newWorksheet,
      ...(await processWorksheet(worksheet.slice(prevIndex, currIndex)))
    ];

    let ws = XLSX.utils.json_to_sheet(newWorksheet);
    workbook.Sheets[sheetName] = ws;
    XLSX.writeFile(workbook, "new_mohave.xlsx");

    progressBar.update(currIndex);

    prevIndex = currIndex;
    if (total - prevIndex < 100) {
      currIndex = total;
    } else {
      currIndex += 100;
    }
  } while (prevIndex < total);

  progressBar.stop();
})();
