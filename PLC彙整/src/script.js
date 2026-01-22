 const ignoreKeywords = [
        'OEE', '標準PPH', '實際PPH', '總故障時間', 
        '總故障次數', '平均排除時間', '每百次故障次數', '標準人機比'
    ];
 const excludeMachines = ['外觀檢測機', '外觀'];


    let sortOrders = {};

    document.getElementById('upload-input').addEventListener('change', function(e) {
        const files = e.target.files;
        const container = document.getElementById('table-container');
        document.getElementById('manual-copy-area').style.display = 'none';
        
        container.innerHTML = '';
        Array.from(files).forEach(file => {
            const reader = new FileReader();
            reader.onload = function(event) {
                const content = event.target.result;
                let data = [];
                if (content.includes('<table') || content.includes('<tr')) {
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(content, 'text/html');
                    const rows = doc.querySelectorAll('tr');
                    rows.forEach(row => {
                        const cells = Array.from(row.querySelectorAll('th, td')).map(td => td.innerText.trim());
                        if (cells.length > 0) data.push(cells);
                    });
                } else {
                    const workbook = XLSX.read(content, { type: 'binary' });
                    data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                }
                if (data.length > 0) renderTable(file.name, data);
            };
            reader.readAsText(file);
        });
    });

    function renderTable(fileName, data) {
        const container = document.getElementById('table-container');
        const tableId = 'table-' + Math.random().toString(36).substr(2, 9);
        const wrapperId = 'wrapper-' + tableId;

        let machineIdx = -1;
        let defectIdx = -1;
        const headers = data[0].map(h => (h || '').toString().trim()); // 預處理標題，去除空白
        
        headers.forEach((h, idx) => {
            if (h.includes('機器名稱')) machineIdx = idx;
            if (h.includes('不良品數') || h.includes('總不良數')) defectIdx = idx;
        });

        let html = `
            <div class="file-header">
                <div class="file-title">📄 ${fileName}</div>
                <div class="btn-group">
                    <button class="btn-copy" onclick="processChartImage('${wrapperId}', '${fileName}')">📷 產生圖表 / 複製</button>
                    <button class="btn-remove" onclick="removeSelected('${tableId}')">🗑 移除選取</button>
                </div>
            </div>
            <div class="table-wrapper" id="${wrapperId}">
                <table id="full-${tableId}">
                    <thead>
                        <tr>
                            <th style="width:30px;" data-html2canvas-ignore>狀態</th>
                            ${headers.map((h, idx) => {
                                const shouldIgnore = ignoreKeywords.some(key => h.includes(key));
                                const ignoreAttr = shouldIgnore ? ' data-html2canvas-ignore' : '';
                                return `<th${ignoreAttr} ondblclick="sortTable('${tableId}', ${idx + 1})">${h}</th>`;
                            }).join('')}
                        </tr>
                    </thead>
                    <tbody id="${tableId}">`;

        for (let i = 1; i < data.length; i++) {
            const rowData = data[i];
            // 排除空行
            if (!rowData || rowData.length === 0 || !rowData.join('').trim()) continue;
            // 排除重複的標題行
            if (rowData[0] === data[0][0]) continue;
            
            let shouldSkip = false;

            // [強化檢查] 判斷機台是否需排除
            if (machineIdx !== -1 && rowData[machineIdx]) {
                const currentMachine = rowData[machineIdx].toString().trim(); // 去除 Excel 儲存格前後空格
                if (excludeMachines.some(key => currentMachine.includes(key))) {
                    shouldSkip = true;
                }
            }

            // 判斷不良數為 0 是否排除 [cite: 50]
            if (defectIdx !== -1 && rowData[defectIdx] !== undefined) {
                let defectVal = parseFloat(rowData[defectIdx].toString().replace(/,/g, ''));
                if (defectVal === 0 || isNaN(defectVal)) shouldSkip = true;
            }

            if (shouldSkip) continue;

            html += `<tr class="data-row">
                <td data-html2canvas-ignore><input type="checkbox" class="row-check"></td>
                ${rowData.map((cell, idx) => {
                    const hName = headers[idx] || '';
                    const shouldIgnore = ignoreKeywords.some(key => hName.includes(key));
                    const ignoreAttr = shouldIgnore ? ' data-html2canvas-ignore' : '';
                    return `<td${ignoreAttr} ondblclick="toggleHighlight(this)">${cell || ''}</td>`;
                }).join('')}
            </tr>`;
        }
        
        html += `</tbody><tfoot id="foot-${tableId}"></tfoot></table></div>`;
        container.insertAdjacentHTML('beforeend', html);
        updateCalculations(tableId);
        Sortable.create(document.getElementById(tableId), { animation: 150 });
    }

    async function processChartImage(elementId, fileName) {
        const element = document.getElementById(elementId);
        if (!element) return;

        const originalBtn = document.querySelector(`button[onclick="processChartImage('${elementId}', '${fileName}')"]`);
        const originalText = originalBtn.innerText;
        originalBtn.innerText = "⏳ 處理中...";
        originalBtn.disabled = true;

        try {
            const canvas = await html2canvas(element, {
                scale:1, // 提高解析度
                backgroundColor: "#ffffff",
                useCORS: true
            });

            canvas.toBlob(async (blob) => {
                try {
                    await navigator.clipboard.write([
                        new ClipboardItem({ 'image/png': blob })
                    ]);
                    alert('✅ 成功！圖表已複製到剪貼簿。');
                } catch (err) {
                    handleFallback(canvas, fileName);
                }
                originalBtn.innerText = originalText;
                originalBtn.disabled = false;
            });
        } catch (error) {
            console.error('截圖錯誤:', error);
            alert('❌ 發生錯誤，請重新整理頁面再試');
            originalBtn.innerText = originalText;
            originalBtn.disabled = false;
        }
    }

    function handleFallback(canvas, fileName) {
        const imgData = canvas.toDataURL("image/png");
        const link = document.createElement('a');
        link.download = `PLC報表_${fileName}.png`;
        link.href = imgData;
        link.click();

        const manualArea = document.getElementById('manual-copy-area');
        const imgTag = document.getElementById('generated-image');
        imgTag.src = imgData;
        manualArea.style.display = 'block';
        manualArea.scrollIntoView({ behavior: 'smooth' });
        alert('⚠️ 注意：無法自動複製。\n\n✅ 已自動下載圖片！\n👇 或在下方圖片「右鍵 -> 複製影像」。');
    }

    window.sortTable = function(tableId, colIndex) {
        const tbody = document.getElementById(tableId);
        const rows = Array.from(tbody.querySelectorAll('tr.data-row'));
        const orderKey = `${tableId}-${colIndex}`;
        sortOrders[orderKey] = sortOrders[orderKey] === 'desc' ? 'asc' : 'desc';
        const isDesc = sortOrders[orderKey] === 'desc';

        rows.sort((a, b) => {
            let valA = a.cells[colIndex].innerText.replace(/,/g, '');
            let valB = b.cells[colIndex].innerText.replace(/,/g, '');
            let numA = parseFloat(valA), numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) return isDesc ? numB - numA : numA - numB;
            return isDesc ? valB.localeCompare(valA, 'zh-Hant') : valA.localeCompare(valB, 'zh-Hant');
        });
        rows.forEach(row => tbody.appendChild(row));
    };

    function removeSelected(tableId) {
        const tbody = document.getElementById(tableId);
        const checks = tbody.querySelectorAll('.row-check:checked');
        if (checks.length === 0) return;
        if (confirm('確定移除選中項目？')) {
            checks.forEach(chk => chk.closest('tr').remove());
            updateCalculations(tableId);
        }
    }

    function updateCalculations(tableId) {
        const tbody = document.getElementById(tableId);
        const tfoot = document.getElementById('foot-' + tableId);
        const rows = tbody.querySelectorAll('tr.data-row');
        const headers = Array.from(document.querySelectorAll(`#full-${tableId} th`));
        const colCount = headers.length;
        let totals = new Array(colCount).fill(0);
        let hasNumber = new Array(colCount).fill(false);
        let isCalculatable = new Array(colCount).fill(true);
        let productionIdx = -1;

        headers.forEach((th, idx) => {
            const title = th.innerText;
            if (title.includes('生產數')) productionIdx = idx;
            const blacklist = ['組別', '代號', 'ID', '率', '%', '料號', '人機比', '標準工', '工時', 'PPH', '時間', '次數'];
            if (blacklist.some(key => title.includes(key))) isCalculatable[idx] = false;
        });

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach((td, idx) => {
                if (idx > 0 && isCalculatable[idx]) {
                    let val = td.innerText.replace(/,/g, '');
                    if (!isNaN(val) && val !== '') {
                        totals[idx] += parseFloat(val);
                        hasNumber[idx] = true;
                    }
                }
            });
        });

        const prodTotal = productionIdx !== -1 ? totals[productionIdx] : 0;
        
        let footHtml = `<tr class="total-row"><td data-html2canvas-ignore>-</td><td>合計</td>`;
        for (let j = 2; j < colCount; j++) {
            const hName = headers[j].innerText;
            const shouldIgnore = ignoreKeywords.some(key => hName.includes(key));
            const ignoreAttr = shouldIgnore ? ' data-html2canvas-ignore' : '';
            let val = (isCalculatable[j] && hasNumber[j]) ? Math.round(totals[j] * 100) / 100 : "";
            footHtml += `<td${ignoreAttr} ondblclick="toggleHighlight(this)">${val}</td>`;
        }
        
        footHtml += `</tr><tr class="ratio-row"><td data-html2canvas-ignore>-</td><td>比率</td>`;
        for (let j = 2; j < colCount; j++) {
            const hName = headers[j].innerText;
            const shouldIgnore = ignoreKeywords.some(key => hName.includes(key));
            const ignoreAttr = shouldIgnore ? ' data-html2canvas-ignore' : '';
            let ratio = (j !== productionIdx && isCalculatable[j] && hasNumber[j] && prodTotal > 0) 
                        ? ((totals[j] / prodTotal) * 100).toFixed(2) + "%" : "";
            footHtml += `<td${ignoreAttr} ondblclick="toggleHighlight(this)">${ratio}</td>`;
        }
        tfoot.innerHTML = footHtml + `</tr>`;
    }

    function toggleHighlight(element) { element.classList.toggle('highlight-red'); }