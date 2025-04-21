let referenceData = null;
let targetData = null;

// 监听文件上传
document.getElementById('referenceFile').addEventListener('change', handleFileUpload);
document.getElementById('targetFile').addEventListener('change', handleFileUpload);
document.getElementById('processBtn').addEventListener('click', processFiles);
document.getElementById('downloadBtn').addEventListener('click', downloadResult);

function handleFileUpload(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        if (event.target.id === 'referenceFile') {
            referenceData = jsonData;
        } else {
            targetData = jsonData;
        }
        
        // 检查是否两个文件都已上传
        if (referenceData && targetData) {
            document.getElementById('processBtn').disabled = false;
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function processFiles() {
    if (!referenceData || !targetData) {
        alert('请先上传两个Excel文件');
        return;
    }

    // 获取参考表的行顺序（使用第一列作为标识）
    const referenceOrder = referenceData.slice(1).map(row => row[0]);
    
    // 创建新的目标数据数组
    const newTargetData = [targetData[0]]; // 保留表头
    
    // 用于跟踪已匹配的行
    const matchedRows = new Set();
    
    // 按照参考顺序重新排列目标数据
    referenceOrder.forEach(key => {
        const matchingRow = targetData.slice(1).find(row => row[0] === key);
        if (matchingRow) {
            newTargetData.push(matchingRow);
            matchedRows.add(matchingRow[0]); // 记录已匹配的行
        }
    });

    // 找出多余的行（未匹配的行）
    const extraRows = targetData.slice(1).filter(row => !matchedRows.has(row[0]));
    
    // 将多余的行添加到末尾
    if (extraRows.length > 0) {
        newTargetData.push(...extraRows);
    }

    // 显示结果
    displayResult(newTargetData, extraRows);
    
    // 保存处理后的数据
    window.processedData = newTargetData;
}

function displayResult(data, extraRows) {
    const resultTable = document.getElementById('resultTable');
    resultTable.innerHTML = '';
    
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    
    // 添加表头
    const headerRow = document.createElement('tr');
    data[0].forEach(cell => {
        const th = document.createElement('th');
        th.textContent = cell;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
    // 添加数据行
    data.slice(1).forEach(row => {
        const tr = document.createElement('tr');
        // 检查是否是多余的行
        const isExtraRow = extraRows.some(extraRow => extraRow[0] === row[0]);
        
        row.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            if (isExtraRow) {
                td.style.backgroundColor = '#fffacd'; // 淡黄色背景
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    
    table.appendChild(thead);
    table.appendChild(tbody);
    resultTable.appendChild(table);
    
    // 显示结果区域
    document.querySelector('.result-section').style.display = 'block';
    
    // 如果有多余的行，显示提示信息
    if (extraRows.length > 0) {
        const infoDiv = document.createElement('div');
        infoDiv.style.marginTop = '10px';
        infoDiv.style.color = '#666';
        infoDiv.textContent = `注意：有 ${extraRows.length} 行数据未在参考文件中找到，已用淡黄色标记并放置在最后。`;
        resultTable.parentNode.insertBefore(infoDiv, resultTable.nextSibling);
    }
}

async function downloadResult() {
    if (!window.processedData) {
        alert('没有可下载的数据');
        return;
    }

    try {
        // 创建工作簿
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet1');

        // 添加数据
        window.processedData.forEach((row, rowIndex) => {
            const excelRow = worksheet.addRow(row);
            
            // 检查是否是多余的行
            const isExtraRow = rowIndex > 0 && !referenceData.slice(1).some(refRow => refRow[0] === row[0]);
            
            if (isExtraRow) {
                // 设置整行的背景色
                excelRow.eachCell((cell) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFACD' }
                    };
                });
            }
        });

        // 设置列宽
        worksheet.columns.forEach(column => {
            column.width = 15;
        });

        // 下载文件
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = '调整后的文件.xlsx';
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error('导出Excel文件时出错:', error);
        alert('导出Excel文件时出错，请重试');
    }
} 