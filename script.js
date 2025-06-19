let analyzedData = null; //儲存後端回傳的json結果
let currentStep = 1; //當前步驟
let totalSteps; //總步驟數
let uploadedPCRFile = null;
let uploadedProductFile = null;

document.addEventListener('DOMContentLoaded', () => { //設定監聽器，等整個html頁面載入完成
    const steps = document.querySelectorAll('.step-content-wrapper'); //每個步驟內容區塊的 DOM 節點
    const stepIndicators = document.querySelectorAll('.step-indicator'); //數字圓圈
    const stepLines = document.querySelectorAll('.step-line'); //步驟進度線
    const prevButton = document.getElementById('prev-button'); //上一步按鈕
    const nextButton = document.getElementById('next-button'); //下一步按鈕

    totalSteps = steps.length; //設定總步驟數

    // 記錄 PCR 和產品 PDF 的上傳檔案
    document.getElementById('regulation-files')?.addEventListener('change', (e) => {
        uploadedPCRFile = e.target.files[0] || null;
    });
    document.getElementById('product-files')?.addEventListener('change', (e) => {
        uploadedProductFile = e.target.files[0] || null;
    });

    function updateStepUI() { //根據所在步驟決定顯示區塊
        steps.forEach((step, index) => {
            step.classList.toggle('active', index + 1 === currentStep);
    });

    stepIndicators.forEach((indicator, index) => { //流程數字圓圈顯示
        const stepNum = index + 1;
        indicator.classList.toggle('active', stepNum === currentStep);
        indicator.classList.toggle('completed', stepNum < currentStep);
        if (stepNum > currentStep) indicator.classList.remove('completed');
    });

    stepLines.forEach((line, index) => { //流程進度線顯示
        const stepNum = index + 1;
        line.classList.toggle('active', stepNum === currentStep - 1 && currentStep > 1);
        line.classList.toggle('completed', stepNum < currentStep - 1);
        if (stepNum >= currentStep - 1) {
            line.classList.remove('completed');
            if (stepNum > currentStep - 1) line.classList.remove('active');
        }
    });

    prevButton.disabled = currentStep === 1; //第一步驟不能點擊上一步
    nextButton.innerHTML = currentStep === totalSteps //調整最後一步的按鈕
    ? '製作下一份碳盤查表 <i class="fas fa-redo"></i>'
    : '下一步 <i class="fas fa-arrow-right"></i>';

    }

    function resetAllData() { //清空所有資料
        analyzedData = [];
        manualInputArea.innerHTML = '';//清空手動填寫的區塊
        if (document.getElementById('regulation-files')) { //清空上傳的PCR檔案
            document.getElementById('regulation-files').value = '';
            document.getElementById('regulation-files-list').innerHTML = '';
        }
        if (document.getElementById('product-files')) { ////清空上傳的產品pdf檔案
            document.getElementById('product-files').value = '';
            document.getElementById('product-files-list').innerHTML = '';
        }
        if (document.getElementById('excel-template')) {
            document.getElementById('excel-template').value = '';
            document.getElementById('excel-template-list').innerHTML = '';
        }

        document.getElementById('computation-status').innerHTML = '';//清空計算與匯出狀態顯示
        document.getElementById('export-status').innerHTML = '';

        const reviewTable = document.getElementById('review-table'); //清空預覽表格
        if (reviewTable) {
            reviewTable.querySelector('thead').innerHTML = '';
            reviewTable.querySelector('tbody').innerHTML = '';
        }
    }

    function handleNext() {
        if (currentStep < totalSteps) {
            if (currentStep > 0 && currentStep - 1 < stepLines.length) {
                stepLines[currentStep - 1].classList.add('completed');
                stepLines[currentStep - 1].classList.remove('active');
            }
            currentStep++;
            updateStepUI();
        } else {
            const confirmed = confirm('此動作將清空所有資料，是否繼續？');
            if (confirmed) {
                resetAllData();       // 使用者按「確定」才清空
                currentStep = 1;
                    // 重新設定步驟狀態樣式
                stepIndicators.forEach((indicator, index) => {
                    indicator.classList.remove('active', 'completed');
                    if (index === 0) indicator.classList.add('active');
                });

                stepLines.forEach((line) => {
                    line.classList.remove('active', 'completed');
                });

                // 重新設定內容區塊樣式
                steps.forEach((step, index) => {
                    step.classList.toggle('active', index === 0);
                });

                // 延遲一個 event loop 再觸發 updateStepUI()
                setTimeout(() => {
                    updateStepUI();
                }, 0);
            }
        }
    }

    function handlePrev() {
        if (currentStep > 1) {
            if (currentStep - 2 >= 0 && currentStep - 2 < stepLines.length) {
                stepLines[currentStep - 2].classList.remove('completed');
                stepLines[currentStep - 2].classList.add('active');
            }
            if (currentStep - 1 < stepLines.length) {
                stepLines[currentStep - 1].classList.remove('active');
            stepLines[currentStep - 1].classList.remove('completed');
            }
            currentStep--;
            updateStepUI();
        }
    }

    nextButton.addEventListener('click', handleNext);
    prevButton.addEventListener('click', handlePrev);

    function setupFileInput(inputId, listId) { //選擇檔案後，把檔案名稱與大小顯示在畫面上
        const fileInput = document.getElementById(inputId);
        const fileListDiv = document.getElementById(listId);
        if (fileInput && fileListDiv) {
        fileInput.addEventListener('change', (event) => {
            fileListDiv.innerHTML = '';
            const files = event.target.files;
            if (files.length > 0) {
                Array.from(files).forEach(file => {
                    const fileItem = document.createElement('div');
                    fileItem.classList.add('file-list-item');
                    fileItem.textContent = `${file.name} (${(file.size / 1024).toFixed(2)} KB)`;
                    fileListDiv.appendChild(fileItem);
            });
            } else {
                fileListDiv.textContent = '未選擇任何文件。';
            }
        });
        }
    }

    setupFileInput('regulation-files', 'regulation-files-list');
    setupFileInput('product-files', 'product-files-list');
    setupFileInput('excel-template', 'excel-template-list');
    // Step 3 上傳檔案並計算
    const computeButton = document.getElementById('compute-button');
    const computationStatus = document.getElementById('computation-status');

    if (computeButton && computationStatus) {
        computeButton.addEventListener('click', async () => {
            if (!uploadedPCRFile || !uploadedProductFile) {
                alert('請確認已上傳 PCR 檔與產品 PDF 檔。');
                return;
            }

            const formData = new FormData();
            formData.append('pcrFile', uploadedPCRFile);
            formData.append('productFile', uploadedProductFile);

            computationStatus.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 系統計算中，請稍候...';

            try {
                const response = await fetch('http://localhost:3000/api/analyze', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();
                if (result.success) {
                    computationStatus.innerHTML = '<i class="fas fa-check-circle"></i> 計算完成！';
                    analyzedData = JSON.parse(result.data.replace(/```json|```/g, '').trim());
                    console.log('分析結果：', analyzedData);
                } else {
                    computationStatus.innerHTML = '錯誤：' + result.error;
                }
            } catch (err) {
                computationStatus.innerHTML = '發送失敗：' + err.message;
            }
        });

    }
// Step 4 手動填寫缺漏欄位
    const manualInputArea = document.getElementById('manual-input-area');
    const saveManualInputButton = document.getElementById('save-manual-input');
    const addNewEntryButton = document.getElementById('add-new-entry');
    const defaultFields = [
        '生命週期階段', '群組', '名稱', '總活動量', '總活動量單位', '每單位數量', '每單位數量單位',
        '名稱', '數值', '排放係數宣告單位', '數據來源', '備註'
    ];
    saveManualInputButton?.addEventListener('click', () => {
        const allInputs = manualInputArea.querySelectorAll('input, select');
        allInputs.forEach(input => {
            const [rowIndex, field] = input.name.split('__');
            if (!analyzedData[rowIndex]) analyzedData[rowIndex] = {};
            analyzedData[rowIndex][field] = input.value.trim();
        });

        console.log('analyzedData 更新結果：', JSON.stringify(analyzedData, null, 2));
        alert('已儲存手動填寫內容！');
    });

    addNewEntryButton?.addEventListener('click', () => {
        if (!analyzedData) analyzedData = [];
        const newEntry = {};
        defaultFields.forEach(field => newEntry[field] = '');
        analyzedData.push(newEntry);
        renderManualInputs();
    });
    function renderManualInputs() { //需重新顯示所有資料列時呼叫
        if (!analyzedData || !Array.isArray(analyzedData)) return; //如果analyedData不存在或非陣列格式就return
        manualInputArea.innerHTML = ''; //清空整個手動填寫區塊
        // 定義下拉選單的選項
        const unitOptions = [
            '毫米(mm)', '公分(cm)', '公尺(m)', '公里(km)', '海浬(nm)', '英寸(in)', '碼(yard)',
            '毫克(mg)', '公克(g)', '公斤(kg)', '公噸(mt)', '英磅(lb)', '毫升(ml)', '公升(L)', '公秉(kl)',
            '平方毫米(mm2)', '平方公分(cm2)', '平方公尺(m2)', '平方公里(km2)',
            '立方毫米(mm3)', '立方公分(cm3)', '立方公尺(m3)', '立方公里(km3)',
            '百萬焦耳(MJ)', '度(kwh)', '延人公里(pkm)', '延噸公里(tkm)',
            'g CO2e', 'kg CO2e', '每平方米‧每小時', '每人‧每小時', '每人', '每人次', '每房-每天',
            '片', '顆', '個', '條', '卷', '瓶', '桶', '盒', '包', '罐', '台', '雙'
        ];

        const stageOptions = [
            '原料取得階段', '製造生產階段', '配銷階段', '使用階段', '廢棄處理階段', '服務階段'
        ];

        const groupOptions = [
            '能源', '資源', '原物料', '輔助項', '產品', '聯產品', '排放', '殘留物'
        ];
        //針對每筆資料row與index畫出對應區塊
        analyzedData.forEach((row, index) => {
            const rowDiv = document.createElement('div'); //建立容器包住這筆資料的所有欄位輸入框
            rowDiv.style.border = '1px solid #ccc';
            rowDiv.style.padding = '12px';
            rowDiv.style.marginBottom = '10px';
            rowDiv.style.borderRadius = '6px';

            const title = document.createElement('strong'); //顯示第幾筆資料
            title.textContent = `第 ${index + 1} 筆資料`;
            rowDiv.appendChild(title);

            const fields = Object.keys(row).length ? Object.keys(row) : defaultFields;//抓出欄位名稱

            fields.forEach((key) => {
                const value = row[key] || '';//抓出每個欄位的值或填空字串
                const div = document.createElement('div');//每個欄位的<div>包裝區域
                div.style.marginTop = '8px';

                const label = document.createElement('label');//加上標籤顯示欄位名稱
                label.textContent = key;
                label.style.marginRight = '8px';

                let input;
                //建立選單或手動填寫格子
                if (key === '生命週期階段') {
                    input = createSelect(stageOptions, value);
                } else if (key === '群組') {
                    input = createSelect(groupOptions, value);
                } else if (
                    key === '總活動量單位' ||
                    key === '每單位數量單位' ||
                    key === '排放係數單位'||
                    key === '排放係數宣告單位'
                ) {
                    input = createSelect(unitOptions, value);
                } else {
                    input = document.createElement('input');
                    input.value = value;
                    input.style.width = '80%';
                    input.style.padding = '6px';
                    input.style.border = '1px solid #ccc';
                    input.style.borderRadius = '4px';
                }

                // 給每個欄位都加 change → 寫入 analyzedData 即時更新
                input.name = `${index}__${key}`;//給每個欄位一個標籤 name = " 列號__欄位名稱 "
                input.addEventListener('change', () => {
                    const [rowIndex, field] = input.name.split('__');
                    if (!analyzedData[rowIndex]) analyzedData[rowIndex] = {};
                    analyzedData[rowIndex][field] = input.value;//當使用者有改值時及時更新到 analyzedData  
                });

                div.appendChild(label); //把先前建立的label(名稱)跟input(麵粉)加到div裡(欄位名+輸入)
                div.appendChild(input);
                rowDiv.appendChild(div);//把先前建立的div加到rowDiv裡(第一列)

            });
            // 建立紅色警示元素
            const warning = document.createElement('div');
            warning.style.color = 'red';
            warning.style.marginTop = '10px';
            warning.style.display = 'none';
            warning.textContent = '單位不一致，請將「總活動量單位」、「每單位數量單位」、「排放係數宣告單位」統一！';
            rowDiv.appendChild(warning);

            //找出三個下拉式選單中選擇的選項，方便後續檢查是否相同
            const selectA = rowDiv.querySelector(`select[name="${index}__總活動量單位"]`); //index指的是第幾筆資料
            const selectB = rowDiv.querySelector(`select[name="${index}__每單位數量單位"]`);
            const selectC = rowDiv.querySelector(`select[name="${index}__排放係數宣告單位"]`);

            function checkUnitsMatch() {
                const a = selectA?.value || '';
                const b = selectB?.value || '';
                const c = selectC?.value || '';
                console.log('🧪 單位比對', { a, b, c });
                const mismatch = a && b && c && (a !== b || a !== c || b !== c); //如果三者有不同則mismatch
                warning.style.display = mismatch ? 'block' : 'none'; //block顯示區塊；none隱藏區塊
            }

            // 綁定 onchange + 寫入即時 updated analyzedData
            [selectA, selectB, selectC].forEach(select => {
                if (!select) return;
                select.addEventListener('change', () => {
                    const [rowIndex, field] = select.name.split('__');
                    if (!analyzedData[rowIndex]) analyzedData[rowIndex] = {};
                    analyzedData[rowIndex][field] = select.value;
                    checkUnitsMatch(); // 每次改值即時檢查
                });
            });

            checkUnitsMatch(); // 初始化執行

            manualInputArea.appendChild(rowDiv);//在manualInputArea這個區塊中加上rowDiv包裝的這筆資料
        });
    }


  // 監控是否進入 step-4，自動渲染欄位
  const observer = new MutationObserver(() => {
    const step4 = document.getElementById('step-4');
    if (step4?.classList.contains('active')) {
        renderManualInputs();
    }
  });

    observer.observe(document.body, { subtree: true, attributes: true });
    
    
    //export Excel
    const exportButton = document.getElementById('export-button');
    const exportStatus = document.getElementById('export-status');

    if (exportButton && exportStatus) {
        exportButton.addEventListener('click', async () => {
            if (!analyzedData || analyzedData.length === 0) {
                alert('請先完成計算步驟，才能匯出報告。');
                return;
            }

        //顯示「匯出中」以及轉圈圈提示
        exportStatus.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 正在產生 Excel 報告...';

        try {
            //送資料到後端api
            const response = await fetch('http://localhost:3000/api/export-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(analyzedData) 
        });

        const result = await response.json();//解析後端回傳的json結果
        if (result.success) {
            exportStatus.innerHTML = '<i class="fas fa-check-circle"></i> Excel 已成功產生！即將下載...';

            // 自動下載 Excel 檔案
            const downloadLink = document.createElement('a'); //建立<a>連結元素
            downloadLink.href = 'http://localhost:3000/generated/Carbon_Footprint_Report.xlsx';

            downloadLink.download = 'Carbon_Footprint_Report.xlsx'; // 設定檔案名稱
            downloadLink.click();  // 模擬點擊下載
        } else {
            exportStatus.innerHTML = '匯出失敗：' + result.error;
        }
        }  catch (err) {
            exportStatus.innerHTML = '發送失敗：' + err.message;
        }
  });
}

    function renderReviewTable() { //預覽表格
    const table = document.getElementById('review-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = '';//重新渲染前先清空舊有資料
    tbody.innerHTML = '';

    if (!analyzedData || !analyzedData.length) {
      tbody.innerHTML = '<tr><td colspan="99">目前尚無資料可預覽。</td></tr>';
      return;
    }

    const headers = Object.keys(analyzedData[0]);//headers抓出資料欄位名稱
    const headerRow = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');//每個欄位都建立一個<th>標題格
      th.textContent = h;
      th.style.borderBottom = '1px solid #aaa';
      th.style.padding = '6px';
      th.style.backgroundColor = '#f0f0f0';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);//將整排標題加到<thead>裡

    analyzedData.forEach(row => {
      const tr = document.createElement('tr');//每筆資料都建立一個<tr>
      headers.forEach(h => {
        const td = document.createElement('td');//每個欄位都建立一個<td>
        td.textContent = row[h] || '';
        td.style.padding = '6px';
        td.style.borderBottom = '1px solid #eee';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  const stepObserver = new MutationObserver(() => {
    const step5 = document.getElementById('step-5');
    if (step5?.classList.contains('active')) {//當切換至step5時就顯示預覽表格
      renderReviewTable();
    }
  });

  stepObserver.observe(document.body, { subtree: true, attributes: true });

function createSelect(options, selectedValue) {//建立select選單
    const select = document.createElement('select');
    select.style.width = '80%';
    select.style.padding = '6px';
    select.style.border = '1px solid #ccc';
    select.style.borderRadius = '4px';

    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '請選擇...';
    if (!selectedValue || selectedValue.trim() === '') emptyOption.selected = true;
    select.appendChild(emptyOption);

    options.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        if (opt === selectedValue) option.selected = true;
        select.appendChild(option);
    });

    return select;
}




  updateStepUI();
});