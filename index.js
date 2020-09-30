const PDFExtract = require("pdf.js-extract").PDFExtract;
const pdfExtract = new PDFExtract();
const fs = require("fs");
const excel = require("exceljs");
let Serialno = 1;

// Created workbook
const workBook = new excel.Workbook();

// Workbook Details
workBook.creator = "Abubakar";
workBook.lastModifiedBy = "Abubakar";
workBook.created = new Date();
workBook.modified = new Date();
workBook.lastPrinted = new Date();
workBook.views = [
  {
    x: 0,
    y: 0,
    width: 10000,
    height: 20000,
    firstSheet: 0,
    activeTab: 1,
    visibility: "visible",
  },
];

// Created worksheet
const sheet = workBook.addWorksheet("SCO Sheet");

// Set sheet colloums
sheet.columns = [
  { header: "Serial No", key: "SrNo", width: 10 },
  { header: "Route No", key: "RouteNo", width: 12 },
  { header: "Account No", key: "AccNo", width: 25 },
  { header: "Application No", key: "AppNo", width: 25 },
  { header: "Date", key: "AppDate", width: 12 },
  { header: "tf", key: "tf", width: 8 },
  { header: "Security Ammount", key: "SecAm", width: 18 },
  { header: "Date Of Payment", key: "DOP", width: 18 },
  { header: "Sco No", key: "ScoNo", width: 16 },
  { header: "Date", key: "ScoDate", width: 12 },
  { header: "Date Of Connection", key: "DOC", width: 18 },
  { header: "Sanctioned Load", key: "SancLoad", width: 16 },
  { header: "Meter no", key: "MeterNo", width: 12 },
  { header: "Name & Address", key: "NameAddress", width: 100 },
  { header: "Mobile No", key: "MobileNo", width: 16 },
  { header: "Month in which feeded", key: "Miwf", width: 25 },
];

// Function to get usefull fields
const getFields = (content, SerialNo) => {
  const outPut = {};
  outPut.SrNo = SerialNo;
  outPut.RouteNo = "";
  outPut.tf = "";
  outPut.DOC = "";
  outPut.MeterNo = "";
  outPut.Miwf = "";
  try {
    for (let field in content) {
      if (content[field].str === "Batch") {
        outPut.Batch = content[Number(field) + 1].str;
      } else if (content[field].str === "Account No.") {
        outPut.AccNo = content[Number(field) + 1].str;
      } else if (content[field].str === "Application No.") {
        outPut.AppNo = content[Number(field) + 1].str;
        outPut.AppDate = content[Number(field) + 3].str;
      } else if (content[field].str === "SCO No.") {
        outPut.ScoNo = content[Number(field) + 1].str;
        outPut.ScoDate = content[Number(field) + 3].str;
      } else if (content[field].str === "Ammount of Security Deposit") {
        outPut.SecAm = content[Number(field) + 1].str;
        outPut.DOP = content[Number(field) + 3].str;
      } else if (content[field].str === "Sanctioned Load") {
        outPut.SancLoad = content[Number(field) + 1].str;
      } else if (content[field].str === "Name & Father/Husband Name") {
        outPut.Name = content[Number(field) + 1].str;
      } else if (content[field].str === "Address") {
        outPut.Address = content[Number(field) + 1].str;
      } else if (content[field].str === "Mobile No.") {
        outPut.MobileNo = content[Number(field) + 1].str;
      }
    }
    outPut.NameAddress = outPut.Name + " / " + outPut.Address;
    delete outPut.Name;
    delete outPut.Address;
  } catch (error) {
    console.log("Error:", error.message);
  }
  return outPut;
};

let NoOfFiles = 0;
let entries = [];
// Gettting list of all files
fs.readdir("./input/", (err, files) => {
  NoOfFiles = files.length;
  files.forEach((file) => {
    // Extracting data from files
    pdfExtract.extract(`./input/${file}`, {}, (err, data) => {
      if (err) return console.log(err);
      entries.push(getFields(data.pages[0].content, Serialno));
      Serialno++;
    });
  });
});

const interval = setInterval(() => {
  // check for tasks
  console.log({ files: NoOfFiles, entr: entries.length });
  if (NoOfFiles === entries.length) {
    entries.forEach((entry) => {
      sheet.addRow(entry);
    });
    // Saving files
    workBook.xlsx
      .writeFile("./output/final.xlsx")
      .then(() => console.log("Done!"))
      .catch((err) => console.log(err.message));
    clearInterval(interval);
  }
}, 1000);
