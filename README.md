<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>QRコードメーカー</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <style>
        body {
            font-family: 'Helvetica Neue', Arial, sans-serif;
            background-color: #f5f7fa;
            color: #333;
            margin: 0;
            padding: 20px;
            line-height: 1.6;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        h1, h2 {
            color: #2c3e50;
            text-align: center;
        }
        .input-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }
        .input-row {
            display: flex;
            flex-direction: column;
            border: 1px solid #ddd;
            border-radius: 5px;
            padding: 10px;
            background-color: #f9f9f9;
        }
        .input-row label {
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
        }
        .input-row input {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-bottom: 8px;
        }
        .add-remove-buttons {
            text-align: center;
            margin: 20px 0;
        }
        .options-area {
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
            flex-wrap: wrap;
            background-color: #f0f4f8;
            padding: 15px;
            border-radius: 5px;
        }
        .option-group {
            margin: 10px;
        }
        button {
            padding: 12px 20px;
            background-color: #3498db;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            transition: background-color 0.3s;
            margin: 5px;
        }
        button:hover {
            background-color: #2980b9;
        }
        button.add-btn {
            background-color: #2ecc71;
        }
        button.add-btn:hover {
            background-color: #27ae60;
        }
        button.remove-btn {
            background-color: #e74c3c;
        }
        button.remove-btn:hover {
            background-color: #c0392b;
        }
        .action-buttons {
            text-align: center;
            margin: 20px 0;
        }
        #results-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        #results-table th, #results-table td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        #results-table th {
            background-color: #f2f2f2;
            position: sticky;
            top: 0;
        }
        .qr-cell {
            text-align: center;
        }
        .results-container {
            margin-top: 30px;
            overflow-x: auto;
        }
        .excel-template {
            margin-top: 20px;
            border: 1px dashed #ccc;
            padding: 15px;
            border-radius: 5px;
            background-color: #f9f9f9;
        }
        .file-input {
            margin: 10px 0;
        }
        .help-text {
            color: #777;
            font-size: 14px;
            margin: 5px 0;
        }
        .status {
            margin: 15px 0;
            padding: 10px;
            border-radius: 5px;
            display: none;
        }
        .success {
            background-color: #d4edda;
            color: #155724;
        }
        .error {
            background-color: #f8d7da;
            color: #721c24;
        }
        .warning {
            background-color: #fff3cd;
            color: #856404;
        }
        .loading {
            text-align: center;
            margin: 20px 0;
            display: none;
        }
        .spinner {
            border: 5px solid #f3f3f3;
            border-top: 5px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 2s linear infinite;
            display: inline-block;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        @media (max-width: 768px) {
            .input-grid {
                grid-template-columns: 1fr;
            }
            .options-area {
                flex-direction: column;
            }
            .container {
                padding: 15px;
            }
        }
        .debug-info {
            font-family: monospace;
            font-size: 12px;
            color: #666;
            margin-top: 5px;
            word-break: break-all;
        }
        .url-length {
            font-size: 12px;
            color: #888;
            margin-top: 2px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>QRコードメーカー</h1>
        
        <div class="options-area">
            <div class="option-group">
                <label>QRコードサイズ:</label>
                <select id="qr-size">
                    <option value="150">小 (150px)</option>
                    <option value="200" selected>中 (200px)</option>
                    <option value="300">大 (300px)</option>
                </select>
            </div>
            
            <div class="option-group">
                <label>エラー訂正レベル:</label>
                <select id="qr-error-level">
                    <option value="L">低 (7%)</option>
                    <option value="M">中 (15%)</option>
                    <option value="Q" selected>高 (25%)</option>
                    <option value="H">最高 (30%)</option>
                </select>
            </div>
            
            <div class="option-group">
                <label>Excelテンプレート:</label>
                <input type="file" id="excel-template" class="file-input" accept=".xlsx,.xls">
                <p class="help-text">テンプレートをアップロードすると、QRコード情報がテンプレートに記入されます</p>
            </div>
        </div>
        
        <div class="add-remove-buttons">
            <button class="add-btn" onclick="addInputRows(5)">入力欄を5つ追加</button>
            <button class="remove-btn" onclick="removeLastInputRows(5)">入力欄を5つ削除</button>
        </div>
        
        <div id="input-container" class="input-grid">
            <!-- 入力欄がここに生成されます -->
        </div>
        
        <div class="action-buttons">
            <button onclick="generateAllQRCodes()">QRコード一括生成</button>
            <button onclick="downloadExcel()">Excelでダウンロード</button>
            <button onclick="downloadAllQRCodes()">全QRコード画像ダウンロード</button>
        </div>
        
        <div id="status" class="status"></div>
        
        <div id="loading" class="loading">
            <div class="spinner"></div>
            <p>処理中...</p>
        </div>
        
        <div class="results-container">
            <h2>生成結果</h2>
            <table id="results-table">
                <thead>
                    <tr>
                        <th>No.</th>
                        <th>QRコード内容</th>
                        <th>説明またはタイトル</th>
                        <th>QRコード</th>
                        <th>QRコードURL</th>
                    </tr>
                </thead>
                <tbody id="results-body">
                    <!-- 結果がここに表示されます -->
                </tbody>
            </table>
        </div>
    </div>

    <script>
        // QRコードのデータを保存する配列
        let qrCodesData = [];
        let templateWorkbook = null;
        let inputRowCount = 0;
        
        // ページ読み込み時に初期入力欄を追加
        window.onload = function() {
            addInputRows(20); // 初期値として20個の入力欄を追加
        };
        
        // 入力欄を追加
        function addInputRows(count) {
            const container = document.getElementById('input-container');
            
            for (let i = 0; i < count; i++) {
                inputRowCount++;
                
                const rowDiv = document.createElement('div');
                rowDiv.className = 'input-row';
                rowDiv.id = `input-row-${inputRowCount}`;
                
                rowDiv.innerHTML = `
                    <label>URL ${inputRowCount}:</label>
                    <input type="text" id="url-${inputRowCount}" placeholder="https://example.com" class="url-input" oninput="updateUrlLength(${inputRowCount})">
                    <div id="url-length-${inputRowCount}" class="url-length"></div>
                    <label>説明/タイトル ${inputRowCount}:</label>
                    <input type="text" id="desc-${inputRowCount}" placeholder="説明またはタイトル（任意）" class="desc-input">
                `;
                
                container.appendChild(rowDiv);
            }
        }
        
        // URLの長さを表示
        function updateUrlLength(rowId) {
            const urlInput = document.getElementById(`url-${rowId}`);
            const urlLengthDiv = document.getElementById(`url-length-${rowId}`);
            
            if (urlInput && urlLengthDiv) {
                const urlLength = urlInput.value.length;
                urlLengthDiv.textContent = `文字数: ${urlLength}文字`;
                
                // 警告表示
                if (urlLength > 500) {
                    urlLengthDiv.style.color = '#e74c3c';
                } else if (urlLength > 300) {
                    urlLengthDiv.style.color = '#f39c12';
                } else {
                    urlLengthDiv.style.color = '#888';
                }
            }
        }
        
        // 入力欄を削除
        function removeLastInputRows(count) {
            const container = document.getElementById('input-container');
            
            for (let i = 0; i < count; i++) {
                if (inputRowCount <= 0) break;
                
                const lastRow = document.getElementById(`input-row-${inputRowCount}`);
                if (lastRow) {
                    container.removeChild(lastRow);
                    inputRowCount--;
                }
            }
        }
        
        // Excelテンプレートの読み込み
        document.getElementById('excel-template').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    templateWorkbook = XLSX.read(data, { type: 'array' });
                    showStatus('テンプレートを読み込みました: ' + file.name, 'success');
                } catch (error) {
                    showStatus('テンプレートの読み込みに失敗しました: ' + error.message, 'error');
                    console.error(error);
                }
            };
            reader.readAsArrayBuffer(file);
        });
        
        // すべてのQRコードを生成
        function generateAllQRCodes() {
            qrCodesData = [];
            let hasData = false;
            
            // 入力欄からデータを収集
            for (let i = 1; i <= inputRowCount; i++) {
                const urlInput = document.getElementById(`url-${i}`);
                const descInput = document.getElementById(`desc-${i}`);
                
                if (urlInput && urlInput.value.trim()) {
                    let url = urlInput.value.trim();
                    const description = descInput ? descInput.value.trim() : '';
                    
                    // URLの簡易検証
                    if (!url.match(/^https?:\/\/.+/i)) {
                        url = 'http://' + url; // プロトコルがなければ追加
                    }
                    
                    qrCodesData.push({
                        number: qrCodesData.length + 1,
                        url: url,
                        description: description,
                        qrCodeDataUrl: null
                    });
                    
                    hasData = true;
                }
            }
            
            if (!hasData) {
                showStatus('URLを少なくとも1つ入力してください', 'error');
                return;
            }
            
            showLoading(true);
            document.getElementById('results-body').innerHTML = '';
            
            // テーブルを生成
            renderTable();
            
            // QRコードを生成
            generateQRCodes().then(() => {
                showStatus('QRコードの生成が完了しました', 'success');
                showLoading(false);
            }).catch(error => {
                showStatus('エラーが発生しました: ' + error.message, 'error');
                showLoading(false);
                console.error(error);
            });
        }
        
        // QRコードを生成する非同期関数（QRiousライブラリ使用）
        async function generateQRCodes() {
            const size = parseInt(document.getElementById('qr-size').value);
            const errorCorrectionLevel = document.getElementById('qr-error-level').value;
            
            // 各QRコードを順番に生成
            for (let i = 0; i < qrCodesData.length; i++) {
                const data = qrCodesData[i];
                
                try {
                    // キャンバス要素を作成
                    const canvas = document.createElement('canvas');
                    canvas.width = size;
                    canvas.height = size;
                    
                    // QRiousを使用してQRコードを生成
                    const qr = new QRious({
                        element: canvas,
                        value: data.url,
                        size: size,
                        level: errorCorrectionLevel
                    });
                    
                    // キャンバスからデータURLを取得
                    data.qrCodeDataUrl = canvas.toDataURL('image/png');
                    
                    // テーブルセルを更新
                    updateQRCodeCell(i);
                    
                    // わずかに遅延させて処理を分散
                    await new Promise(resolve => setTimeout(resolve, 10));
                    
                } catch (error) {
                    console.error(`QRコード #${i+1} 生成エラー:`, error);
                    showStatus(`QRコード #${i+1} の生成中にエラーが発生しました: ${error.message}`, 'warning');
                    
                    // エラーが発生してもすべてを停止せず、次に進む
                    data.qrCodeDataUrl = null;
                    updateQRCodeErrorCell(i, error);
                }
            }
        }
        
        // テーブルを初期レンダリング
        function renderTable() {
            const tbody = document.getElementById('results-body');
            tbody.innerHTML = '';
            
            qrCodesData.forEach((data, index) => {
                const row = document.createElement('tr');
                
                row.innerHTML = `
                    <td>${data.number}</td>
                    <td>${escapeHtml(data.url)}<div class="debug-info">長さ: ${data.url.length}文字</div></td>
                    <td>${escapeHtml(data.description)}</td>
                    <td class="qr-cell" id="qr-cell-${index}"><div class="spinner"></div></td>
                    <td>${escapeHtml(data.url)}</td>
                `;
                
                tbody.appendChild(row);
            });
        }
        
        // QRコードセルを更新
        function updateQRCodeCell(index) {
            const cell = document.getElementById(`qr-cell-${index}`);
            if (!cell) return;
            
            const data = qrCodesData[index];
            if (data.qrCodeDataUrl) {
                cell.innerHTML = `
                    <img src="${data.qrCodeDataUrl}" alt="QR Code" style="max-width: 200px;">
                    <br>
                    <a href="${data.qrCodeDataUrl}" download="qrcode-${index + 1}.png" class="download-link">ダウンロード</a>
                `;
            }
        }
        
        // QRコードエラーセルを更新
        function updateQRCodeErrorCell(index, error) {
            const cell = document.getElementById(`qr-cell-${index}`);
            if (!cell) return;
            
            cell.innerHTML = `
                <div style="color: #e74c3c;">エラー: QRコードを生成できませんでした</div>
                <div class="debug-info">${escapeHtml(error.message)}</div>
            `;
        }
        
        // Excelでダウンロード
        function downloadExcel() {
            if (qrCodesData.length === 0) {
                showStatus('ダウンロードするデータがありません。先にQRコードを生成してください。', 'error');
                return;
            }
            
            showLoading(true);
            
            try {
                let workbook;
                
                if (templateWorkbook) {
                    // テンプレートがある場合はそれを使用
                    workbook = XLSX.utils.book_new();
                    // シートをコピー
                    for (let i = 0; i < templateWorkbook.SheetNames.length; i++) {
                        const sheetName = templateWorkbook.SheetNames[i];
                        const worksheet = templateWorkbook.Sheets[sheetName];
                        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                    }
                    
                    // 最初のシートを取得
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // データをテンプレートに書き込む
                    qrCodesData.forEach((data, index) => {
                        // 行が1から始まるため、ヘッダー行を考慮して+2する
                        const rowIndex = index + 2;
                        
                        // A列: QRコード内容
                        worksheet[`A${rowIndex}`] = { t: 's', v: data.url };
                        
                        // B列: 説明またはタイトル
                        worksheet[`B${rowIndex}`] = { t: 's', v: data.description || '' };
                        
                        // D列: QRコードURL
                        worksheet[`D${rowIndex}`] = { t: 's', v: data.url };
                    });
                } else {
                    // 新しいワークブックを作成
                    workbook = XLSX.utils.book_new();
                    
                    // データを準備
                    const wsData = [
                        ['QRコード内容', '説明またはタイトル', 'QRコード', 'QRコードURL'],
                        ...qrCodesData.map(data => [data.url, data.description || '', '(QRコード画像)', data.url])
                    ];
                    
                    // ワークシートを作成
                    const worksheet = XLSX.utils.aoa_to_sheet(wsData);
                    
                    // ワークブックに追加
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'QRコード一覧');
                }
                
                // Excelファイルを生成してダウンロード
                XLSX.writeFile(workbook, 'qr_codes.xlsx');
                
                showStatus('Excelファイルをダウンロードしました', 'success');
            } catch (error) {
                showStatus('Excelファイルの生成に失敗しました: ' + error.message, 'error');
                console.error(error);
            }
            
            showLoading(false);
        }
        
        // すべてのQRコード画像をダウンロード
        function downloadAllQRCodes() {
            const validQRCodes = qrCodesData.filter(data => data.qrCodeDataUrl);
            
            if (validQRCodes.length === 0) {
                showStatus('ダウンロードするQRコードがありません。先にQRコードを生成してください。', 'error');
                return;
            }
            
            showStatus(`${validQRCodes.length}個のQRコード画像のダウンロードを開始します...`, 'success');
            
            // ダウンロード処理を少し遅延させてUIを更新
            setTimeout(() => {
                // すべてのQRコード画像を個別にダウンロード
                validQRCodes.forEach((data, index) => {
                    const link = document.createElement('a');
                    link.href = data.qrCodeDataUrl;
                    link.download = `qrcode-${data.number}.png`;
                    document.body.appendChild(link);
                    
                    // 少し遅延させてダウンロードを開始
                    setTimeout(() => {
                        link.click();
                        document.body.removeChild(link);
                    }, 100 * index);
                });
            }, 500);
        }
        
        // ステータスメッセージを表示
        function showStatus(message, type) {
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = message;
            statusDiv.className = 'status ' + type;
            statusDiv.style.display = 'block';
            
            // 5秒後に非表示（エラーの場合は10秒）
            const timeout = type === 'error' ? 10000 : 5000;
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, timeout);
        }
        
        // ローディング表示を切り替え
        function showLoading(show) {
            document.getElementById('loading').style.display = show ? 'block' : 'none';
        }
        
        // HTMLをエスケープ
        function escapeHtml(text) {
            if (!text) return '';
            return text
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");
        }
    </script>
</body>
</html>
