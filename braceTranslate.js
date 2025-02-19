const XLSX = require('xlsx');

const file1Path = '02_MSTag.xlsx'; 
const batterPath = 'find_new_key.xlsx'; 

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    return data;
}

// 读取两个 Excel 文件
const MSTag = readExcel(file1Path).slice(3); // 去除前三行
const batterDate = readExcel(batterPath);

// 检查 Translate 是否包含 MSTag 的 Simp_TIMI，并添加 newTranslate 字段
function checkAndFormatTranslate(item) {
  const translate = item.Translate;

  // 确保 translate 是一个字符串
  if (typeof translate !== 'string') {
    item.newTranslate = ''; // 如果 translate 不是字符串，设置 newTranslate 为空字符串
    return item;
  }

  // 复制原始的 translate 字符串，用于替换
  let newTranslate = translate;

  // 遍历 MSTag 数组
  let hasMatch = false;
  MSTag.forEach(tag => {
    const regex = new RegExp(tag.Simp_TIMI, 'g');
    if (regex.test(newTranslate)) {
      hasMatch = true;
      newTranslate = newTranslate.replace(regex, `{${tag.Simp_TIMI}}`);
    }
  });

  // 如果没有匹配，则将 newTranslate 设置为空字符串
  if (!hasMatch) {
    item.newTranslate = '';
  } else {
    item.newTranslate = newTranslate;
  }

  return item;
}

// 遍历 batterDate 数组并应用 checkAndFormatTranslate 函数
const newBatterDate = batterDate.map(checkAndFormatTranslate);

// console.log(newBatterDate);

// 将 JSON 数据转换为工作表
const worksheet = XLSX.utils.json_to_sheet(newBatterDate);

// 创建一个新的工作簿并添加工作表
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// 写入 Excel 文件
XLSX.writeFile(workbook, 'braceTranslate20250229.xlsx');