document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const verifyBtn = document.getElementById('verifyBtn');
    const loading = document.getElementById('loading');
    const resultsSection = document.getElementById('resultsSection');
    const resultsList = document.getElementById('resultsList');
    const aiToggle = document.getElementById('aiToggle');
    const aiText = document.getElementById('aiText');
    const apiKeyInput = document.getElementById('apiKey');

    // Drag & Drop
    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        fileName.textContent = file.name;
        verifyBtn.disabled = false;
        resultsSection.classList.add('hidden');
    }

    verifyBtn.addEventListener('click', async () => {
        const file = fileInput.files[0] || (fileName.textContent ? { name: fileName.textContent } : null);
        if (!fileInput.files[0]) return;

        verifyBtn.disabled = true;
        loading.classList.remove('hidden');
        resultsSection.classList.add('hidden');

        const apiKey = apiKeyInput.value.trim();

        if (apiKey) {
            aiText.textContent = "AI (Gemini) 正在分析文件 (PDF/圖片)...";
        } else if (aiToggle.checked) {
            aiText.textContent = "AI 正在進行 Excel 格式辨識與轉換...";
        } else {
            aiText.textContent = "";
        }

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('use_ai', aiToggle.checked);
        formData.append('api_key', apiKey);

        try {
            const response = await fetch('/verify', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (response.ok) {
                displayResults(data.results);
            } else {
                alert('上傳失敗: ' + (data.error || 'Unknown error'));
            }
        } catch (error) {
            alert('發生錯誤: ' + error);
        } finally {
            loading.classList.add('hidden');
            verifyBtn.disabled = false;
        }
    });

    function displayResults(results) {
        resultsList.innerHTML = '';

        if (!results || results.length === 0) {
            resultsList.innerHTML = '<div class="result-item"><p>未發現可檢核的數據，或數據完全符合規則。</p></div>';
            return;
        }

        let counts = { OK: 0, WARNING: 0, ERROR: 0 };

        results.forEach(res => {
            counts[res.status] = (counts[res.status] || 0) + 1;

            const item = document.createElement('div');
            item.className = `result-item ${res.status}`;
            item.innerHTML = `
                <div class="result-content">
                    <span class="category">${res.category}</span>
                    <div style="display:flex; align-items:center;">
                        <span class="tag ${res.status}">${res.status}</span>
                        <div>
                            <h4>${res.rule}</h4>
                            <p>${res.message}</p>
                        </div>
                    </div>
                </div>
            `;
            resultsList.appendChild(item);
        });

        document.getElementById('totalChecks').textContent = results.length;
        document.getElementById('totalErrors').textContent = counts.ERROR;
        document.getElementById('totalWarnings').textContent = counts.WARNING;
        document.getElementById('totalOk').textContent = counts.OK;

        resultsSection.classList.remove('hidden');
    }
});
