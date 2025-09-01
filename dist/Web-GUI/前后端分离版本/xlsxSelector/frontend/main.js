// 全局文件列表
let mergeFiles = [];
let dedupeRefFiles = [];

// 后端 API 地址，确保与 Flask 应用的端口一致
const API_BASE_URL = 'http://127.0.0.1:5000';

function showTab(tabId) {
    document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');
    document.querySelectorAll('.tab-button').forEach(button => button.classList.remove('active'));
    document.querySelector(`.tab-button[onclick="showTab('${tabId}')"]`).classList.add('active');
}

// 分割方式切换逻辑
document.querySelectorAll('input[name="method"]').forEach(radio => {
    radio.addEventListener('change', (event) => {
        document.querySelector('.split-count').style.display = 'none';
        document.querySelector('.split-range').style.display = 'none';
        
        const selectedMethod = event.target.value;
        if (selectedMethod === 'count') {
            document.querySelector('.split-count').style.display = 'block';
        } else if (selectedMethod === 'range') {
            document.querySelector('.split-range').style.display = 'block';
        }
    });
});

// 模糊查重阈值切换
function toggleFuzzyThreshold() {
    const container = document.getElementById('fuzzy-threshold-container');
    const mode = document.querySelector('input[name="mode"]:checked').value;
    if (mode === 'fuzzy') {
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}

// 动态获取列名和行数
async function getColumns(file, columnsContainerId, targetInputId, isMultiSelect, infoElementId) {
    const columnsContainer = document.getElementById(columnsContainerId);
    const targetInput = document.getElementById(targetInputId);
    const infoElement = document.getElementById(infoElementId);
    
    if (!file) {
        columnsContainer.innerHTML = '';
        if (infoElement) infoElement.textContent = '';
        return;
    }

    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const response = await fetch(`${API_BASE_URL}/api/get_file_info`, {
            method: 'POST',
            body: formData,
        });
        const data = await response.json();
        
        columnsContainer.innerHTML = '';
        if (data.columns) {
            columnsContainer.style.display = 'block';
            if (infoElement) infoElement.textContent = `总行数: ${data.row_count}`;
            data.columns.forEach(column => {
                const button = document.createElement('span');
                button.className = 'column-button';
                button.textContent = column;
                button.onclick = () => {
                    if (isMultiSelect) {
                        const currentCols = targetInput.value.split(',').map(c => c.trim()).filter(c => c);
                        const isSelected = currentCols.includes(column);
                        
                        if (isSelected) {
                            targetInput.value = currentCols.filter(c => c !== column).join(',');
                        } else {
                            targetInput.value = [...currentCols, column].join(',');
                        }
                        button.classList.toggle('selected', !isSelected);
                    } else {
                        targetInput.value = column;
                        columnsContainer.querySelectorAll('.column-button').forEach(btn => btn.classList.remove('selected'));
                        button.classList.add('selected');
                    }
                    updateRenameInputs(targetInput.value, columnsContainerId.replace('columns', 'rename-cols'));
                };
                columnsContainer.appendChild(button);
            });
        } else {
            columnsContainer.innerHTML = `<p style="color:red;font-size:14px;">${data.error}</p>`;
        }
    } catch (error) {
        console.error('Error fetching file info:', error);
        columnsContainer.innerHTML = `<p style="color:red;font-size:14px;">获取文件信息失败。</p>`;
    }
}
window.getColumns = getColumns;

// 更新重命名输入框
function updateRenameInputs(originalCols, renameInputId) {
    const renameInput = document.getElementById(renameInputId);
    const colsArray = originalCols.split(',').map(c => c.trim()).filter(c => c);
    if (renameInput) {
        const newRenameValues = colsArray.map(col => {
            const currentValues = renameInput.value.split(',').map(c => c.trim()).filter(c => c);
            const existingIndex = currentValues.indexOf(col);
            return existingIndex !== -1 ? currentValues[existingIndex] : col;
        });
        renameInput.value = newRenameValues.join(',');
    }
}

// 文件累加和显示逻辑
function handleFileSelection(fileInputId, fileListId, fileArray, onFileChangeCallback = null) {
    const fileInput = document.getElementById(fileInputId);
    const fileList = document.getElementById(fileListId);
    
    fileInput.addEventListener('change', (event) => {
        const newFiles = Array.from(event.target.files);
        if (newFiles.length === 0) return;
        
        fileArray.push(...newFiles);
        updateFileList(fileList, fileArray, onFileChangeCallback);
        
        event.target.value = '';
    });

    function updateFileList(listContainer, array, callback) {
        listContainer.innerHTML = '';
        if (array.length === 0) {
            listContainer.innerHTML = '<p style="color:#6c757d;font-style:italic;">暂无文件</p>';
            if (callback) callback(array);
            return;
        }
        
        array.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            if (fileInputId === 'dedupe-ref-input') {
                 fileItem.innerHTML = `
                    <div>
                        <span>${file.name}</span>
                        <span class="remove-file" data-index="${index}">✖</span>
                        <div id="dedupe-ref-file-cols-${index}" class="column-selection"></div>
                        <input type="text" name="ref_col_${index}" id="dedupe_ref_col_input_${index}" required>
                    </div>
                `;
                listContainer.appendChild(fileItem);
                
                getColumns(file, `dedupe-ref-file-cols-${index}`, `dedupe_ref_col_input_${index}`, false);
            } else {
                fileItem.innerHTML = `
                    <span>${file.name}</span>
                    <span class="remove-file" data-index="${index}">✖</span>
                `;
                listContainer.appendChild(fileItem);
            }
        });
        if (callback) callback(array);
    }
    
    fileList.addEventListener('click', (event) => {
        if (event.target.classList.contains('remove-file')) {
            const index = parseInt(event.target.getAttribute('data-index'));
            fileArray.splice(index, 1);
            updateFileList(fileList, fileArray, onFileChangeCallback);
        }
    });
    
    updateFileList(fileList, fileArray, onFileChangeCallback);
}

// 统一处理表单提交
document.querySelectorAll('form').forEach(form => {
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const statusDiv = document.getElementById('status-message');
        statusDiv.className = 'alert alert-success';
        statusDiv.textContent = '文件正在处理中，请稍候...';
        
        const formData = new FormData(form);
        const apiPath = form.id === 'merge-form' ? '/api/merge' :
                        form.id === 'split-form' ? '/api/split' :
                        form.id === 'dedupe-form' ? '/api/deduplicate' :
                        '/api/clean';

        if (form.id === 'merge-form') {
            mergeFiles.forEach(file => {
                formData.append('files', file);
            });
        } else if (form.id === 'dedupe-form') {
            formData.append('main_file', document.querySelector('input[name="main_file_input"]').files[0]);
            dedupeRefFiles.forEach((file, index) => {
                formData.append('ref_files', file);
                formData.append(`ref_col_${index}`, document.getElementById(`dedupe_ref_col_input_${index}`).value);
            });
        }
        
        try {
            const response = await fetch(`${API_BASE_URL}${apiPath}`, {
                method: 'POST',
                body: formData,
            });

            if (response.ok) {
                const blob = await response.blob();
                const disposition = response.headers.get('Content-Disposition');
                let filename = 'download';
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) {
                        filename = decodeURIComponent(matches[1].replace(/['"]/g, ''));
                    }
                }

                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                statusDiv.className = 'alert alert-success';
                statusDiv.textContent = '✅ 处理成功，文件已开始下载。';
            } else {
                const errorText = await response.text();
                statusDiv.className = 'alert alert-danger';
                statusDiv.textContent = `❌ 处理失败: ${errorText}`;
            }
        } catch (error) {
            console.error('Fetch error:', error);
            statusDiv.className = 'alert alert-danger';
            statusDiv.textContent = `❌ 请求失败，请检查网络或服务器。`;
        }
    });
});

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    handleFileSelection('merge-file-input', 'merge-file-list', mergeFiles, (files) => {
        const combinedColumns = new Set();
        const promises = files.map(file => {
            const formData = new FormData();
            formData.append('file', file);
            return fetch(`${API_BASE_URL}/api/get_file_info`, { method: 'POST', body: formData })
                .then(res => res.json())
                .then(data => {
                    if (data.columns) {
                        data.columns.forEach(col => combinedColumns.add(col));
                    }
                });
        });
        Promise.all(promises).then(() => {
            const columnsContainer = document.getElementById('merge-columns');
            const targetInput = document.getElementById('merge_cols_input');
            columnsContainer.innerHTML = '';
            if (combinedColumns.size > 0) {
                columnsContainer.style.display = 'block';
                Array.from(combinedColumns).sort().forEach(column => {
                    const button = document.createElement('span');
                    button.className = 'column-button';
                    button.textContent = column;
                    button.onclick = () => {
                        const currentCols = targetInput.value.split(',').map(c => c.trim()).filter(c => c);
                        const isSelected = currentCols.includes(column);
                        if (isSelected) {
                            targetInput.value = currentCols.filter(c => c !== column).join(',');
                        } else {
                            targetInput.value = [...currentCols, column].join(',');
                        }
                        button.classList.toggle('selected', !isSelected);
                        updateRenameInputs(targetInput.value, 'merge_rename_cols_input');
                    };
                    columnsContainer.appendChild(button);
                });
            } else {
                columnsContainer.innerHTML = '<p style="color:red;font-size:14px;">未找到共同列名或文件无效</p>';
                columnsContainer.style.display = 'block';
            }
        });
    });
    handleFileSelection('dedupe-ref-input', 'dedupe-ref-list', dedupeRefFiles);
    toggleFuzzyThreshold();
});