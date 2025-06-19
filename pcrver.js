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
import path from 'path';




const __dirname = dirname(fileURLToPath(import.meta.url));

dotenv.config();

  // 用於處理 Excel 文件

const genAI = new GoogleGenerativeAI(process.env.API_KEY);
const fileManager = new GoogleAIFileManager(process.env.API_KEY);


export async function uploadPDF(pcrFiles, productFiles){
  const carbonCSVContent = await fs.readFile(`${__dirname}/Preview_Data (1).csv`, 'utf-8');
  const prompt=await fs.readFile(`${__dirname}/填表教學.txt`,'utf-8');
  const pull=await fs.readFile(`${__dirname}/下拉式.txt`,'utf-8');
  const group=await fs.readFile(`${__dirname}/群組定義.txt`,'utf-8');
  const model = genAI.getGenerativeModel({
    model: 'gemini-1.5-pro',
    
    generationConfig: {
      temperature: 0,
      topP: 1,
      topK: 1
    }
});
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
  // 第一階段：上傳 PCR 與 PDF，讓 Gemini 擷取已提供資料
   const result1 = await model.generateContent([
  {
    fileData: { fileUri: pcruploadResult.file.uri, mimeType: pcruploadResult.file.mimeType }
  },
  {
    fileData: { fileUri: useruploadResult.file.uri, mimeType: useruploadResult.file.mimeType }
  },
`請根據我提供的碳盤查報告 PDF，找出本報告中記載的「本年度總產品產量」或「生產數量」。  

請直接以 JSON 格式回傳，格式如下：

{
  "總產品產量": 數值,
  "單位": "原文單位"
}
`])
const extractedJSON0 = result1.response.text();
  const cleanJsonString0 = extractedJSON0
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();
console.log(cleanJsonString0);

const result = await model.generateContent([
  {
    fileData: { fileUri: pcruploadResult.file.uri, mimeType: pcruploadResult.file.mimeType }
  },
  {
    fileData: { fileUri: useruploadResult.file.uri, mimeType: useruploadResult.file.mimeType }
  },
  `${pull}`,
  `${cleanJsonString0}`,
  `${group}`,
  `
我提供了PCR 文件 ${pcruploadResult}，請先完整閱讀並理解如何進行產品碳盤查。接著閱讀我提供的碳盤查報告 PDF ${useruploadResult}，以及我剛才提供的下拉選單 ${pull}。

請根據以上所有內容，「以 JSON 格式產出一份填寫完整的碳盤查表格」，這份表格將被用來填入 Excel 表單，因此必須嚴格對應下列欄位與格式，且只允許填入碳盤查報告 PDF 中明確提供的資訊。禁止從其他來源補齊或推測。

請僅回傳 JSON 格式，不得包含任何解釋文字或補充說明。

✅ 表格欄位結構如下（對應 Excel）：
Excel 格式 JSON Key 名稱
A2 "生命週期階段"
B2 "群組"
C2 "名稱"
D2 "總活動量"
E2 "總活動量單位"
F2 "每單位數量"
G2 "每單位數量單位"
H2 "排放係數名稱"
I2 "排放係數數值"
J2 "排放係數單位"
K2 "排放係數數據來源"
L2 "備註"

❗ 必須遵守的規則如下：
📌 在開始提取資料前，請先從碳盤查報告 PDF 中找出「本年度產品總產量」或「總生產數量」，這個數量將用於計算所有活動項目的總活動量。

請務必遵守以下邏輯：

1. 每筆活動數據中：
  總產量${cleanJsonString0}
- 「總活動量」=「每單位數量」×「產品總產量」
- 禁止直接將每單位數量填入總活動量欄位
- 總活動量必須是最終數值，禁止寫成算式（如 1030 × 2530 ❌）


3. 「排放係數數值」「排放係數單位」「排放係數數據來源」三欄，只能填寫盤查報告 PDF 中明確出現的內容。禁止引用其他來源（如資料庫、外部網站、你自己的知識）來補齊這些欄位，尤其「排放係數數據來源」的部分應更加注意請勿填入任何使用者未說明的來源。
  🚫 即使你知道某些資料，也禁止填寫。禁止任何情況下使用模型記憶、語料庫、常識、過往學習資料補填本次表格中的排放係數欄位。  
  ⚠️ 注意：「排放係數數據來源（K 欄）」預設為空。只有在碳盤查報告 PDF 中明確出現來源說明（如：「數據來源：環境部產品碳足跡資訊網」）時，才可以填入該來源。
  即使你知道此數值出現在外部資料庫，若報告未明載來源資訊，請將「數據來源」欄保持空白。

4. 若報告中未提供該數據，禁止使用「無」、「未知」、「尚未提供」等文字，留空即可。

 5. **每單位數量（F 欄）** 指的是：當一個產品的功能單位為 1（如 1 公噸），活動數據的總量除以總產品數量所得的每單位值，請直接給出換算後的數值，例如：
        -  正確：1030
        -  錯誤：1030/11（不能出現算式）
        如果報告中沒有明寫每單位數量，請用總活動量除以產品總產量來計算每單位數量。

  6. **總活動量（D 欄）** 必須根據「每單位數量 * 總產量或總重量」直接算出結果，不能出現計算式，請直接填好。例如：
        -  正確：11330
        -  錯誤：1030 * 11
      總活動量須謹慎注意要瑱入的是「標的產品總活動量」，請謹慎判斷資料提供的是「全廠所有產品總活動量」還是「標的產品總活動量」，若是全廠所有產品總活動量則需要資料中有提供「標的產品佔全廠產品比例」，並由該比例乘上全廠所有產品總活動量才能得到我要的「標的產品總活動量」。
      若有提供總活動量以及每單位數量但兩者之間並不符合「每單位數量*總產量=總活動量」，則請以每單位數量為主並重新計算總活動量。

 7. 注意 "總活動量" 與 "每單位數量" 的差別，以及注意功能單位可能寫在報告裡的某一處。

8. 「排放係數單位」需為功能單位（例如「公斤 CO₂e／立方公尺」→「立方公尺」），並與下拉選單選項一致。

9. 若報告中僅提供排放係數數值，未提及來源，請僅填入數值，來源請保持空白。若某一筆資料中，報告未提供排放係數資訊（數值、單位、來源），請保持這三個欄位為空。若有提及排放係數數值，則也請務必填寫排放係數名稱欄位。

10. 所有選項（生命週期階段、群組、單位）皆必須嚴格從 ${pull} 提供的下拉選單中選取，嚴格禁止任何自創或修改，例如「公斤」則需選取選單中對應的「公斤(kg)」。
  排放係數的單位若有寫「CO₂e／」後面接上單位，則只需填上後面的單位即可。例如：「CO₂e／公斤」則填上「公斤(kg)」，「CO₂e／延噸公里」填上「延噸公里(tkm)」，但如果單位只有「CO₂e」，那就填上「CO₂e」就好。
  （生命週期階段、群組、單位）中填入的資料都必須非常嚴格遵守與下拉選單檔案中的選項一模一樣，否則會有嚴重錯誤。

11.群組欄位之定義如${group}所示

🔍 掃描與擷取原則：
請完整掃描整份 Qcake 碳盤查報告與 PCR 文件，逐頁逐段檢查所有與活動數據有關的內容。

活動數據是指任何導致碳排的活動，例如：

原物料與加工品

能源（電、蒸汽、燃料等）

包裝材料與包裝廢棄物

冷媒、製冷設備、水資源

廢棄物處理

運輸與物流配送

銷售、儲存與產品後處理

即使僅在報告某處簡單提及，也必須提取為一筆資料。

請將上述所有資訊轉換為 JSON 陣列格式的資料表，每列對應一筆活動。請勿包含任何額外說明、註解或標頭文字，亦不需包含總產量資料，請嚴格遵守我提供的json格式。


  `
]);

const extractedJSON = result.response.text();
  const cleanJsonString1 = extractedJSON
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();
console.log(cleanJsonString1);
// 第二階段：補齊缺漏的排放係數
const filledResult = await model.generateContent([
  `${pull}`,
  `${carbonCSVContent}`,
  `
以下是從報告中擷取出的 JSON 資料：
${cleanJsonString1}

請根據以下規則，**僅補齊空白的排放係數欄位**，並嚴格禁止任何覆寫已填寫過的資料：

🔒【嚴格禁止修改原資料】
1. 若某筆資料中「排放係數數值」(I欄) 已經有內容，請**絕對禁止更動下列欄位**，不論是否符合資料庫內容，也不論數值是否正確：
   - "排放係數名稱" (H欄)
   - "排放係數數值" (I欄)
   - "排放係數單位" (J欄)
   - "排放係數數據來源" (K欄)
   - ⚠️ 不允許任何情況下覆蓋這四欄，即使資料庫中有看似更合適的內容。
   - ⚠️ 在「排放係數數值」(I欄) 已經有內容的情況下，即使"排放係數數據來源"為空欄也嚴格禁止填入資料。

✅【允許補齊空欄】
2. 僅當「排放係數數值」(I欄) 為空時，才允許從我提供的排放係數資料庫（${carbonCSVContent}）中進行補齊。補齊條件如下：
   - 根據欄位「名稱」（C欄）對應查找資料庫中的「項目名稱」
   - 僅選用排放係數表中「宣告單位」與 JSON 中「每單位數量單位」(G欄) 完全一致的資料
   - 若找不到單位一致的資料，則該筆資料的排放係數三欄保持空白

📌 補齊內容應包含：
   - "排放係數名稱" → 對應資料庫中名稱
   - "排放係數數值" → 資料庫中最新年度的數值
   - "排放係數單位" → 僅可選擇 ${pull} 中提供的功能單位
   - "排放係數數據來源" → 使用格式：「環境部,2023」（對應資料庫的 department 與 announcementyear）

⛔ 禁止任何自創或推測行為，禁止根據語意「猜你認為合適的」補齊，若不符合條件，請保留空白。

⛔ 不可調整或格式化原有資料，即使你覺得結果比較整齊或一致，也請保留原始格式不動。

請僅回傳 JSON 陣列格式，不得包含任何解釋文字、註解、說明或提示。

✅ 表格欄位如下：
A2 "生命週期階段"
B2 "群組"
C2 "名稱"
D2 "總活動量"
E2 "總活動量單位"
F2 "每單位數量"
G2 "每單位數量單位"
H2 "排放係數名稱"
I2 "排放係數數值"
J2 "排放係數單位"
K2 "排放係數數據來源"
L2 "備註"
`
]);


const finalJSON = filledResult.response.text();

  const cleanJsonString = finalJSON
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();

  return cleanJsonString;
}


async function generateExcel(datain) {
  let data=datain;

    if (!Array.isArray(data)) {
      throw new Error('傳入的資料不是陣列');
    }
    console.log("收到 Excel 資料：", data);
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
  
    await workbook.xlsx.writeFile(path.join(__dirname, 'generated', 'Carbon_Footprint_Report.xlsx'));

    console.log('Excel 文件含選單已生成');
  }
  
  

async function main(){
    
    const text=await uploadPDF();
    await generateExcel(text);
}

//main();

export { generateExcel };