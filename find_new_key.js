const XLSX = require('xlsx');

// 文件路径定义，注意路径中的双反斜杠
const oldFilePath = `C:\\Users\\v_jinlqi\\Desktop\\compare_excel\\braceTranslate20250110.xlsx`;
const newFilePath = 'C:\\Users\\v_jinlqi\\Desktop\\compare_excel\\Merged_five_266.xlsx';

// 读取 Excel 文件
function readExcel(filePath, fileName) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet).map(item => ({ ...item, "来源": fileName }));

    return data;
}

// 读取2个 Excel 文件
const oldFileData = readExcel(oldFilePath, "oldFilePath");
const newFileData = readExcel(newFilePath, "newFilePath");

console.log(oldFileData);
console.log(newFileData);