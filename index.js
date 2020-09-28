// const PDFExtract = require("pdf.js-extract").PDFExtract;
// const pdfExtract = new PDFExtract();

const xlsx = require("xlsx");
const wb = xlsx.readFile("blank sco statement.xls");
const rows = xlsx.utils.sheet_to_json(wb.Sheets["Sheet1"]);

// for(let pdf of pdffiles){
//     const rowFields = {};
//     pdfExtract.extract("./Service_Cost_Order.pdf", {} /* options*/, (err, data) => {
//       if (err) return console.log(err);
//       const fields = data.pages[0].content;
//       for (let field in fields) {
//         if (fields[field].str === "Batch") {
//           console.log("Batch : ", fields[Number(field) + 1].str);
//         } else if (fields[field].str === "Account No.") {
//           console.log("Acc no : ", fields[Number(field) + 1].str);
//         } else if (fields[field].str === "Application No.") {
//           console.log("Application No : ", fields[Number(field) + 1].str);
//         } else if (fields[field].str === "SCO No.") {
//           console.log("SCO No : ", fields[Number(field) + 1].str);
//         }
//       }
//     });
// }

const newWb = xlsx.utils.book_new();
const newWs = xlsx.utils.json_to_sheet(rows);
xlsx.utils.book_append_sheet(newWb, newWs, "Sheet1");

xlsx.writeFile(newWb, "sco_sheet.xls");
