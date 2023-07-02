const fs = require("fs")
const XLSX = require("xlsx")

const workbook = XLSX.readFile("./dado.xlsx")

let workssheets = {}
for (const sheetname of workbook.SheetNames){
    workssheets[sheetname] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetname])
}
console.log(workssheets)