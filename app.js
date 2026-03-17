// 获取 DOM 元素
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const convertBtn = document.getElementById('convertBtn');
const status = document.getElementById('status');

let selectedFile = null;

// 点击上传区域触发文件选择
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// 拖拽上传
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

// 文件选择
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

// 处理文件
function handleFile(file) {
    // 检查文件类型
    if (!file.name.endsWith('.docx')) {
        showStatus('请选择 .docx 格式的文件', 'error');
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.classList.add('show');
    status.classList.remove('show');
}

// 转换按钮点击
convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    showStatus('正在转换中...', 'processing');
    convertBtn.disabled = true;

    try {
        // 读取 Word 文件
        const arrayBuffer = await readFileAsArrayBuffer(selectedFile);

        // 使用 mammoth 转换为 HTML（保留表格结构）
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;

        // 将 HTML 转换为 Excel
        convertHtmlToExcel(html, selectedFile.name);

    } catch (error) {
        console.error('转换失败:', error);
        showStatus('转换失败，请重试', 'error');
        convertBtn.disabled = false;
    }
});

// 读取文件为 ArrayBuffer
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(e);
        reader.readAsArrayBuffer(file);
    });
}

// 将 HTML 转换为 Excel（保留表格）
function convertHtmlToExcel(html, originalFileName) {
    // 创建一个临时 div 来解析 HTML
    const div = document.createElement('div');
    div.innerHTML = html;

    // 查找所有表格
    const tables = div.querySelectorAll('table');
    const workbook = XLSX.utils.book_new();

    if (tables.length === 0) {
        // 如果没有表格，提取纯文本
        const text = div.innerText || div.textContent;
        const lines = text.split('\n').filter(line => line.trim() !== '');
        const data = lines.map((line, index) => [index + 1, line]);

        const worksheet = XLSX.utils.aoa_to_sheet([['序号', '内容'], ...data]);
        worksheet['!cols'] = [
            { wch: 8 },
            { wch: 100 }
        ];
        XLSX.utils.book_append_sheet(workbook, worksheet, '文档内容');
    } else {
        // 如果有表格，逐个处理
        tables.forEach((table, index) => {
            const rows = table.querySelectorAll('tr');
            const data = [];

            rows.forEach(row => {
                const cells = row.querySelectorAll('td, th');
                const rowData = [];
                cells.forEach(cell => {
                    // 清理单元格内容，去除多余空白
                    let text = cell.innerText || cell.textContent || '';
                    text = text.replace(/\s+/g, ' ').trim();
                    rowData.push(text);
                });
                if (rowData.length > 0) {
                    data.push(rowData);
                }
            });

            if (data.length > 0) {
                const worksheet = XLSX.utils.aoa_to_sheet(data);

                // 自动设置列宽
                const colWidths = [];
                data.forEach(row => {
                    row.forEach((cell, colIndex) => {
                        const width = Math.min(Math.max(cell.length * 2, 10), 50);
                        colWidths[colIndex] = Math.max(colWidths[colIndex] || 0, width);
                    });
                });
                worksheet['!cols'] = colWidths.map(w => ({ wch: w }));

                // 设置表头样式（加粗）
                if (data.length > 0) {
                    for (let col = 0; col < data[0].length; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                        if (!worksheet[cellAddress]) continue;
                        worksheet[cellAddress].s = {
                            font: { bold: true }
                        };
                    }
                }

                const sheetName = `表格${index + 1}`;
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            }
        });
    }

    // 生成 Excel 文件名
    const excelFileName = originalFileName.replace('.docx', '.xlsx');

    // 下载文件
    XLSX.writeFile(workbook, excelFileName);

    showStatus('转换成功！文件已开始下载', 'success');
    convertBtn.disabled = false;
}

// 显示状态
function showStatus(message, type) {
    status.textContent = message;
    status.className = 'status show ' + type;

    // 成功消息 3 秒后隐藏
    if (type === 'success') {
        setTimeout(() => {
            status.classList.remove('show');
        }, 3000);
    }
}