let analyzedData = null;
let currentStep = 1;
let totalSteps;
let uploadedPCRFile = null;
let uploadedProductFile = null;

document.addEventListener('DOMContentLoaded', () => {
  const steps = document.querySelectorAll('.step-content-wrapper');
  const stepIndicators = document.querySelectorAll('.step-indicator');
  const stepLines = document.querySelectorAll('.step-line');
  const prevButton = document.getElementById('prev-button');
  const nextButton = document.getElementById('next-button');

  
  totalSteps = steps.length;

    // ✅ 記錄 PCR 和產品 PDF 的上傳檔案
    document.getElementById('regulation-files')?.addEventListener('change', (e) => {
    uploadedPCRFile = e.target.files[0] || null;
    });

    document.getElementById('product-files')?.addEventListener('change', (e) => {
    uploadedProductFile = e.target.files[0] || null;
    });

  function updateStepUI() {
    steps.forEach((step, index) => {
      step.classList.toggle('active', index + 1 === currentStep);
    });

    stepIndicators.forEach((indicator, index) => {
      const stepNum = index + 1;
      indicator.classList.toggle('active', stepNum === currentStep);
      indicator.classList.toggle('completed', stepNum < currentStep);
      if (stepNum > currentStep) indicator.classList.remove('completed');
    });

    stepLines.forEach((line, index) => {
      const stepNum = index + 1;
      line.classList.toggle('active', stepNum === currentStep - 1 && currentStep > 1);
      line.classList.toggle('completed', stepNum < currentStep - 1);
      if (stepNum >= currentStep - 1) {
        line.classList.remove('completed');
        if (stepNum > currentStep - 1) line.classList.remove('active');
      }
    });

    prevButton.disabled = currentStep === 1;
    nextButton.innerHTML = currentStep === totalSteps 
    ? '製作下一份碳盤查表 <i class="fas fa-redo"></i>'
    : '下一步 <i class="fas fa-arrow-right"></i>';

  }
function resetAllData() {
  analyzedData = [];
  manualInputArea.innerHTML = '';
    if (document.getElementById('regulation-files')) {
    document.getElementById('regulation-files').value = '';
    document.getElementById('regulation-files-list').innerHTML = '';
    }
    if (document.getElementById('product-files')) {
    document.getElementById('product-files').value = '';
    document.getElementById('product-files-list').innerHTML = '';
    }
    if (document.getElementById('excel-template')) {
    document.getElementById('excel-template').value = '';
    document.getElementById('excel-template-list').innerHTML = '';
    }

  document.getElementById('computation-status').innerHTML = '';
  document.getElementById('export-status').innerHTML = '';
  const reviewTable = document.getElementById('review-table');
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
            resetAllData();       // ✅ 使用者按「確定」才清空
            currentStep = 1;
                  // ✅ 重新設定步驟狀態樣式
            stepIndicators.forEach((indicator, index) => {
                indicator.classList.remove('active', 'completed');
                if (index === 0) indicator.classList.add('active');
            });

            stepLines.forEach((line) => {
                line.classList.remove('active', 'completed');
            });

            // ✅ 重新設定內容區塊樣式
            steps.forEach((step, index) => {
                step.classList.toggle('active', index === 0);
            });

            // 🔁 延遲一個 event loop 再觸發 updateStepUI()
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

  function setupFileInput(inputId, listId) {
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
      console.log('✅ 分析結果：', analyzedData);
    } else {
      computationStatus.innerHTML = '❌ 錯誤：' + result.error;
    }
  } catch (err) {
    computationStatus.innerHTML = '❌ 發送失敗：' + err.message;
  }
});

  }
// ✅ Step 4 手動填寫缺漏欄位
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

    console.log('✅ analyzedData 更新結果：', JSON.stringify(analyzedData, null, 2));
    alert('✅ 已儲存手動填寫內容！');
    });

    addNewEntryButton?.addEventListener('click', () => {
        if (!analyzedData) analyzedData = [];
        const newEntry = {};
        defaultFields.forEach(field => newEntry[field] = '');
        analyzedData.push(newEntry);
        renderManualInputs();
    });
  function renderManualInputs() {
  if (!analyzedData || !Array.isArray(analyzedData)) return;
  manualInputArea.innerHTML = '';

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

  analyzedData.forEach((row, index) => {
    const rowDiv = document.createElement('div');
    rowDiv.style.border = '1px solid #ccc';
    rowDiv.style.padding = '12px';
    rowDiv.style.marginBottom = '10px';
    rowDiv.style.borderRadius = '6px';

    const title = document.createElement('strong');
    title.textContent = `第 ${index + 1} 筆資料`;
    rowDiv.appendChild(title);

    const fields = Object.keys(row).length ? Object.keys(row) : defaultFields;

    fields.forEach((key) => {
      const value = row[key] || '';
      const div = document.createElement('div');
      div.style.marginTop = '8px';

      const label = document.createElement('label');
      label.textContent = key;
      label.style.marginRight = '8px';

      let input;

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
    input.name = `${index}__${key}`;
    input.addEventListener('change', () => {
    const [rowIndex, field] = input.name.split('__');
    if (!analyzedData[rowIndex]) analyzedData[rowIndex] = {};
    analyzedData[rowIndex][field] = input.value;
    });

      div.appendChild(label);
      div.appendChild(input);
      rowDiv.appendChild(div);

    });
    // 建立紅色警示元素
    const warning = document.createElement('div');
    warning.style.color = 'red';
    warning.style.marginTop = '10px';
    warning.style.display = 'none';
    warning.textContent = '❌ 單位不一致，請將「總活動量單位」、「每單位數量單位」、「排放係數宣告單位」統一！';
    rowDiv.appendChild(warning);

    // 抓取 select
    const selectA = rowDiv.querySelector(`select[name="${index}__總活動量單位"]`);
    const selectB = rowDiv.querySelector(`select[name="${index}__每單位數量單位"]`);
    const selectC = rowDiv.querySelector(`select[name="${index}__排放係數宣告單位"]`);

    function checkUnitsMatch() {
    const a = selectA?.value || '';
    const b = selectB?.value || '';
    const c = selectC?.value || '';
    const mismatch = a && b && c && (a !== b || a !== c || b !== c);
    warning.style.display = mismatch ? 'block' : 'none';
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

    manualInputArea.appendChild(rowDiv);
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
  const exportButton = document.getElementById('export-button');
  const exportStatus = document.getElementById('export-status');

  if (exportButton && exportStatus) {
    exportButton.addEventListener('click', async () => {
      if (!analyzedData) {
        alert('請先完成計算步驟，才能匯出報告。');
        return;
      }

      exportStatus.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 正在產生 Excel 報告...';

      try {
        const response = await fetch('http://localhost:3000/api/export-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(analyzedData)
        });

        const result = await response.json();
        if (result.success) {
          exportStatus.innerHTML = '<i class="fas fa-check-circle"></i> Excel 已成功產生！請至伺服器端查看檔案。';
        } else {
          exportStatus.innerHTML = '❌ 匯出失敗：' + result.error;
        }
      } catch (err) {
        exportStatus.innerHTML = '❌ 發送失敗：' + err.message;
      }
    });
  }
    function renderReviewTable() {
    const table = document.getElementById('review-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!analyzedData || !analyzedData.length) {
      tbody.innerHTML = '<tr><td colspan="99">目前尚無資料可預覽。</td></tr>';
      return;
    }

    const headers = Object.keys(analyzedData[0]);
    const headerRow = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      th.style.borderBottom = '1px solid #aaa';
      th.style.padding = '6px';
      th.style.backgroundColor = '#f0f0f0';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    analyzedData.forEach(row => {
      const tr = document.createElement('tr');
      headers.forEach(h => {
        const td = document.createElement('td');
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
    if (step5?.classList.contains('active')) {
      renderReviewTable();
    }
  });

  stepObserver.observe(document.body, { subtree: true, attributes: true });

function createSelect(options, selectedValue) {
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
