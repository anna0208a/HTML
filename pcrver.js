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

export async function uploadPDF(filepath){
    const carbonCSVContent = await fs.readFile(`${__dirname}/Preview_Data (1).csv`, 'utf-8');
    const prompt=await fs.readFile(`${__dirname}/填表教學.txt`,'utf-8');
    const pull=await fs.readFile(`${__dirname}/下拉式.txt`,'utf-8');
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-pro" });
    const pcruploadResult = await fileManager.uploadFile(
    `${__dirname}/烘焙蛋糕pcr.pdf`,
    {
        mimeType: "application/pdf",
    },
    );
    const useruploadResult = await fileManager.uploadFile(
    filepath,
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
        `我提供了多份pcr文件，請先閱讀完所有pcr文件${pcruploadResult}學習如何進行碳盤查後，閱讀我提供的Qcake碳盤查報告.pdf${useruploadResult}、以及我剛才教你的填表方式，「以json格式完成一份填寫完整的碳盤查表格，只要回覆我這個表格就好不要有其他任何解釋」。此表格必須可以用來填入excel並且嚴格遵守下列格式:A1到G1為一合併儲存格，內容為"活動數據"。H1到K1為一合併儲存格，內容為"排放係數"。A1到G1為一合併儲存格，內容為"活動數據"。L1、L2為一合併儲存格，內容為"備註"。A2"生命週期階段"。B2"群組"。C2"名稱"。D2"總活動量"。E2"總活動量單位"。F2"每單位數量"。G2"每單位數量單位"。H2"名稱"。I2"數值"。J2"排放係數宣告單位"。K2"數據來源"。json的key須以A2到K2來命名。其中H2、I2、J2、K2的欄位皆在形容使用的排放係數，排放係數由我提供的數據${carbonCSVContent}取得，csv檔案中的department那行代表數據來源。表中要填入的排放係數以及排放係數單位實際代表的是碳足跡數值（kgCO2e）以及宣告單位。排放係數的單位必須要嚴格遵守與總活動量單位以及每單位數量單位相同且數值也必須進行單位換算(例如如果結果是6公噸但單位須轉為公斤時，需轉成6000公斤)!!在資料中找不到的格子就直接留白就好不要補充說明。另外，當中的生命週期階段欄位、群組欄位以及各個單位欄位都是下拉式選單，選單內容請嚴格依照${pull}的內容，選項不能有任何更動。另外總活動量數據為每單位活動量乘上總產量或總重量(依單位決定)，請直接寫出計算後的結果絕對不要列算式，「每單位」的部分也請幫我算好1個單位的數量，不要列類似8/200這種形式的`
    ]); 
    const resultText = await result.response.text();
     const result_missing = await model.generateContent([
        { fileData: { fileUri: pcruploadResult.file.uri, mimeType: pcruploadResult.file.mimeType }},
        { fileData: { fileUri: useruploadResult.file.uri, mimeType: useruploadResult.file.mimeType }},
        `請先閱讀完所有pcr文件${pcruploadResult}以及填表教學${prompt}學習如何進行碳盤查後，閱讀目前為止的結果${resultText}，條列式告訴我我還缺少哪些「活動數據名稱項目」(即整列都沒有抓取到的，如果只缺部分欄位的項目就不算)，並不要有其他解釋。使用者只需在每個需要填寫的大方向都有提供類似項目即可，只需提供明確缺少的大方向項目。`
      ]);
      const missingInfo = result_missing.response.text(); // 顯示用
      console.log("缺少的資訊:", missingInfo); // 顯示用 
    console.log(result.response.text());
    const resultText2 = await result.response.text();
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
        item['名稱'],
        item['數值'],
        item['排放係數宣告單位'],
        item['數據來源'],
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
