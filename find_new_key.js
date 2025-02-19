const XLSX = require('xlsx');

// 文件路径定义，注意使用双反斜杠避免转义字符问题
const oldFilePath = `C:\\Users\\v_jinlqi\\Desktop\\compare_excel\\braceTranslate20250110.xlsx`;
const newFilePath = 'C:\\Users\\v_jinlqi\\Desktop\\compare_excel\\Merged_five_266.xlsx';

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    return data;
}

// 读取两个 Excel 文件
const oldFileData = readExcel(oldFilePath);
const newFileData = readExcel(newFilePath);

// 提取旧文件中的所有Key
const oldKeys = new Set(oldFileData.map(item => item.Key));

// 找出新文件中新增的行
const newRows = newFileData.filter(item => !oldKeys.has(item.Key));

console.log("新增的行：", newRows);

// 将 JSON 数据转换为工作表
const worksheet = XLSX.utils.json_to_sheet(newRows);

// 创建一个新的工作簿并添加工作表
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// 写入 Excel 文件
XLSX.writeFile(workbook, 'find_new_key.xlsx')