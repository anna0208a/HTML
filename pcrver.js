import { GoogleAIFileManager } from "@google/generative-ai/server";
import { GoogleGenerativeAI } from "@google/generative-ai";
import fs from 'fs/promises';
import ExcelJS from "exceljs";
import pdf2table from 'pdf2table';
import axios from 'axios';
import { resolve } from "path";
import https from 'https';
import dotenv from 'dotenv';
import { dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

dotenv.config();

  // 用於處理 Excel 文件

const genAI = new GoogleGenerativeAI(process.env.API_KEY);
const fileManager = new GoogleAIFileManager(process.env.API_KEY);

/* async function PDFtoTable(pdfuri){
    const response = await axios.get(pdfuri, {
        responseType: 'arraybuffer',
        headers:{
            'Authorization':`Bearer ${process.env.API_KEY}`,
        }});
    return new Promise((resolve,reject)=>{
        pdf2table.parse(response.data,function(err,rows){
            resolve(rows);
        });
    
});
} */

export async function uploadPDF(pcrFiles, productFiles){
    const carbonCSVContent = await fs.readFile(`${__dirname}/Preview_Data (1).csv`, 'utf-8');
    const prompt=await fs.readFile(`${__dirname}/填表教學.txt`,'utf-8');
    const pull=await fs.readFile(`${__dirname}/下拉式.txt`,'utf-8');
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-pro" });
    const pcruploadResult = await fileManager.uploadFile(
      pcrFiles,
      {
        mimeType: "application/pdf",
      },
    );
    const useruploadResult = await fileManager.uploadFile(
    productFiles,
    {
        mimeType: "application/pdf",
    },
    );

    const exampleCSVContent = await fs.readFile(`${__dirname}/盤查項目範本.csv`, 'utf-8');
    const result3 = await model.generateContent([
        `${prompt}`,
    ]);
    const result = await model.generateContent([
        {
            fileData: {
                fileUri: pcruploadResult.file.uri,
                mimeType: pcruploadResult.file.mimeType,
            },
            fileData:{
                fileUri: useruploadResult.file.uri,
                mimeType: useruploadResult.file.mimeType,
            },

        },
        `我提供了多份 PCR 文件 ${pcruploadResult}，請先完整閱讀並理解如何進行產品碳盤查。接著閱讀我提供的 Qcake 碳盤查報告 PDF ${useruploadResult}，以及我剛才提供的填表教學方式與下拉選單 ${pull}。

請根據以上所有內容，「以 JSON 格式產出一份填寫完整的碳盤查表格」，這份表格會被用來填入 Excel 表單，因此必須**嚴格對應下列欄位與格式**，且「只回傳 JSON 格式，請勿包含任何解釋文字」。

---

### ✅ 表格欄位結構如下（對應 Excel）：

- A2: "生命週期階段"
- B2: "群組"
- C2: "名稱"
- D2: "總活動量"
- E2: "總活動量單位"
- F2: "每單位數量"
- G2: "每單位數量單位"
- H2: "排放係數名稱"
- I2: "排放係數數值"
- J2: "排放係數單位"
- K2: "排放係數數據來源"
- L2: "備註"

請將這些欄位名直接當作 JSON 的 Key 名稱。

---

### ❗ 特別注意事項（請務必遵守）：

1. **每單位數量（F 欄）** 指的是：當一個產品的功能單位為 1（如 1 公噸），活動數據的總量除以總產品數量所得的每單位值，請直接給出換算後的數值，例如：
   - ✅ 正確：1030
   - ❌ 錯誤：1030/11（不能出現算式）

2. **總活動量（D 欄）** 必須根據「每單位數量 × 總產量或總重量」直接算出結果，不能出現計算式，請直接填好。例如：
   - ✅ 正確：11330
   - ❌ 錯誤：1030 × 11

3. **不得預設每單位數量為 1。** 這是一個常見錯誤。請務必根據報告中實際提供的總活動量與總產量（或重量）做換算。

4. **排放係數與其單位（I、J 欄）** 必須從我提供的 ${carbonCSVContent} 中取得，並依據排放係數名稱對應查找。若排放係數單位與活動量單位不同，**請務必進行單位換算後再填入**，例如：6 噸需換為 6000 公斤。

5. 若找不到欄位對應數據，請直接空白處理，不要填「無」或「未知」等補充說明。

6. 「生命週期階段」、「群組」、「單位」的選項必須嚴格從 ${pull} 提供的下拉選單中選取，不得自行創造或更動任何選項。

---

請輸出這份 JSON 表格。
`
    ]); 
    const resultText2 = await result.response.text();
    const result_missing = await model.generateContent([
        { fileData: { fileUri: pcruploadResult.file.uri, mimeType: pcruploadResult.file.mimeType }},
        { fileData: { fileUri: useruploadResult.file.uri, mimeType: useruploadResult.file.mimeType }},
        `請先閱讀完所有pcr文件${pcruploadResult}以及填表教學${prompt}學習如何進行碳盤查後，閱讀目前為止的結果${resultText2}，條列式告訴我我還缺少哪些「活動數據名稱項目」(即整列都沒有抓取到的，如果只缺部分欄位的項目就不算)，並不要有其他解釋。使用者只需在每個需要填寫的大方向都有提供類似項目即可，只需提供明確缺少的大方向項目。`
      ]);
      const missingInfo = result_missing.response.text(); // 顯示用
      console.log("缺少的資訊:", missingInfo); // 顯示用
    console.log(result.response.text());
const cleanJsonString = resultText2
  .replace(/```json/g, '')
  .replace(/```/g, '')
  .trim();

return cleanJsonString;
}


async function generateExcel(datain) {
  let data=datain;
/*     const cleanJsonString = datain
      .replace(/```json/g, '')
      .replace(/```/g, '')
      .trim();
  
    let data = JSON.parse(cleanJsonString); */
    if (!Array.isArray(data)) {
      throw new Error('❌ 傳入的資料不是陣列');
    }
    console.log("📦 收到 Excel 資料：", data);
    const workbook = new ExcelJS.Workbook();
    const ws_main = workbook.addWorksheet('Sheet1');
    const ws_dropdown = workbook.addWorksheet('Code');
  
    // 下拉選單內容（生命週期階段、群組、單位）
    const dropdowns = {
      "生命週期階段": [
        '原料取得階段', '製造生產階段', '配銷階段', '使用階段', '廢棄處理階段', '服務階段'
      ],
      "群組": [
        '能源', '資源', '原物料', '輔助項', '產品', '聯產品', '排放', '殘留物'
      ],
      "單位": [
        '毫米(mm)', '公分(cm)', '公尺(m)', '公里(km)', '海浬(nm)', '英寸(in)', '碼(yard)',
        '毫克(mg)', '公克(g)', '公斤(kg)', '公噸(mt)', '英磅(lb)', '毫升(ml)', '公升(L)', '公秉(kl)',
        '平方毫米(mm2)', '平方公分(cm2)', '平方公尺(m2)', '平方公里(km2)', '立方毫米(mm3)', '立方公分(cm3)',
        '立方公尺(m3)', '立方公里(km3)', '百萬焦耳(MJ)', '度(kwh)', '延人公里(pkm)', '延噸公里(tkm)',
        'g CO2e', 'kg CO2e', '每平方米‧每小時', '每人‧每小時', '每人', '每人次', '每房-每天',
        '片', '顆', '個', '條', '卷', '瓶', '桶', '盒', '包', '罐', '台', '雙'
      ]
    };
  
    // 寫入選單內容到工作表 Code
    dropdowns["生命週期階段"].forEach((v, i) => ws_dropdown.getCell(`A${i + 1}`).value = v);
    dropdowns["群組"].forEach((v, i) => ws_dropdown.getCell(`C${i + 1}`).value = v);
    dropdowns["單位"].forEach((v, i) => ws_dropdown.getCell(`E${i + 1}`).value = v);
    dropdowns["單位"].forEach((v, i) => ws_dropdown.getCell(`G${i + 1}`).value = v);
    dropdowns["單位"].forEach((v, i) => ws_dropdown.getCell(`J${i + 1}`).value = v);

    // 主表標題設置
    ws_main.mergeCells('A1:G1');
    ws_main.getCell('A1').value = '活動數據';
    ws_main.mergeCells('H1:K1');
    ws_main.getCell('H1').value = '排放係數';
  
    const headers = [
      '生命週期階段', '群組', '名稱', '總活動量', '單位', '每單位數量', '單位',
      '排放名稱', '數值', '單位', '數據來源', '備註'
    ];
    ws_main.addRow(headers);
    ws_main.mergeCells('L1:L2');
    ws_main.getCell('L1').value = '備註';
    // 寫入資料內容
    data.forEach(item => {
      ws_main.addRow([
        item['生命週期階段'],
        item['群組'],
        item['名稱'],
        item['總活動量'],
        item['總活動量單位'],
        item['每單位數量'],
        item['每單位數量單位'],
        item['排放係數名稱'],
        item['排放係數數值'],
        item['排放係數單位'],
        item['排放係數數據來源'],
        item['備註'] || ''
      ]);
    });
  
    // 建立資料驗證公式
    const unitFormula = `Code!$G$1:$G$${dropdowns["單位"].length}`;
    const lifecycleFormula = `Code!$A$1:$A$${dropdowns["生命週期階段"].length}`;
    const groupFormula = `Code!$C$1:$C$${dropdowns["群組"].length}`;
  
    const rowStart = 3;
    const rowEnd = data.length + 2;
    for (let r = rowStart; r <= rowEnd; r++) {
      ws_main.getCell(`A${r}`).dataValidation = {
        type: 'list', formulae: [lifecycleFormula], allowBlank: true
      };
      ws_main.getCell(`B${r}`).dataValidation = {
        type: 'list', formulae: [groupFormula], allowBlank: true
      };
      ['E', 'G', 'J'].forEach(col => {
        ws_main.getCell(`${col}${r}`).dataValidation = {
          type: 'list', formulae: [unitFormula], allowBlank: true
        };
      });
    }
  
    await workbook.xlsx.writeFile('Carbon_Footprint_Report.xlsx');
    console.log('✅ Excel 文件含選單已生成');
  }
  
  

async function main(){
    
    const text=await uploadPDF();
    await generateExcel(text);
}

//main();

export { generateExcel };
