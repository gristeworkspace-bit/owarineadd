// ============================================
// Constants
// ============================================
const CORS_PROXY = 'https://corsproxy.io/?';
const YAHOO_API_BASE = 'https://query1.finance.yahoo.com/v8/finance/chart/';
const BATCH_SIZE = 5;
const BATCH_DELAY_MS = 1500;

// ============================================
// State
// ============================================
let parsedData = null;      // { metadata: string[], header: string[], rows: string[][] }
let closingPrices = {};     // { stockCode: { price: number|null, dividend: number|null, date: string } }
let errorMessages = [];
let isFetching = false;

// ============================================
// DOM References
// ============================================
const $ = (id) => document.getElementById(id);

const dom = {
    dropZone: $('dropZone'),
    fileInput: $('fileInput'),
    fileInfo: $('fileInfo'),
    fileName: $('fileName'),
    fileSize: $('fileSize'),
    fileRemove: $('fileRemove'),
    settingsSection: $('settingsSection'),
    dateColumn: $('dateColumn'),
    stockCount: $('stockCount'),
    uniqueStockCount: $('uniqueStockCount'),
    actionSection: $('actionSection'),
    fetchBtn: $('fetchBtn'),
    downloadBtn: $('downloadBtn'),
    progressContainer: $('progressContainer'),
    progressLabel: $('progressLabel'),
    progressValue: $('progressValue'),
    progressFill: $('progressFill'),
    progressDetail: $('progressDetail'),
    resultsSection: $('resultsSection'),
    successCount: $('successCount'),
    naCount: $('naCount'),
    errorCount: $('errorCount'),
    tableSection: $('tableSection'),
    rowCount: $('rowCount'),
    tableHead: $('tableHead'),
    tableBody: $('tableBody'),
    errorSection: $('errorSection'),
    errorLog: $('errorLog'),
};

// ============================================
// CSV Parser
// ============================================

/**
 * マルチライン引用符フィールドに対応したCSVパーサー
 */
function parseCSVRows(text) {
    const rows = [];
    let currentRow = [];
    let currentField = '';
    let inQuote = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];

        if (inQuote) {
            if (ch === '"') {
                if (i + 1 < text.length && text[i + 1] === '"') {
                    currentField += '"';
                    i++;
                } else {
                    inQuote = false;
                }
            } else {
                currentField += ch;
            }
        } else {
            if (ch === '"') {
                inQuote = true;
            } else if (ch === ',') {
                currentRow.push(currentField);
                currentField = '';
            } else if (ch === '\r') {
                // skip
            } else if (ch === '\n') {
                currentRow.push(currentField);
                currentField = '';
                rows.push(currentRow);
                currentRow = [];
            } else {
                currentField += ch;
            }
        }
    }

    // 最後の行
    if (currentField !== '' || currentRow.length > 0) {
        currentRow.push(currentField);
        rows.push(currentRow);
    }

    return rows.filter(row => row.some(f => f.trim() !== ''));
}

/**
 * CSVテキスト全体を解析
 * ヘッダー3行（メタデータ）+ マルチラインヘッダー + データ行
 */
function parseCSV(text) {
    // 全行を取得
    const allLines = text.split(/\r?\n/);

    // メタデータ行を特定（ヘッダー行は「基準日」で始まる）
    let headerStartIdx = -1;
    for (let i = 0; i < Math.min(10, allLines.length); i++) {
        if (allLines[i].startsWith('基準日')) {
            headerStartIdx = i;
            break;
        }
    }

    if (headerStartIdx === -1) {
        throw new Error('CSVのヘッダー行が見つかりません。「基準日」で始まる行が必要です。');
    }

    const metadata = allLines.slice(0, headerStartIdx);

    // ヘッダー行以降をマルチラインCSVとして解析
    const remaining = allLines.slice(headerStartIdx).join('\n');
    const allRows = parseCSVRows(remaining);

    if (allRows.length < 2) {
        throw new Error('データ行が見つかりません。');
    }

    // ヘッダー行のフィールド名を整理（改行を除去）
    const rawHeader = allRows[0];
    const header = rawHeader.map(h => h.replace(/\n/g, ''));

    // データ行
    const rows = allRows.slice(1);

    return { metadata, header, rows };
}

// ============================================
// Stock Code Utilities
// ============================================

/**
 * 銘柄コード列から Yahoo Finance 用ティッカーを生成
 * 例: "82270 " → "8227.T"
 */
function toTicker(rawCode) {
    const code = rawCode.trim();
    if (code.length < 4) return null;

    // 先頭4文字を取得
    const ticker = code.substring(0, 4);

    // 英数字のみで構成されているか確認
    if (!/^[A-Za-z0-9]+$/.test(ticker)) return null;

    return ticker + '.T';
}

/**
 * データからユニークな銘柄コード一覧を取得
 */
function getUniqueStocks(rows, codeColIdx, dateColIdx) {
    const seen = new Map(); // stockCode → { ticker, date, rawCode }

    for (const row of rows) {
        if (row.length <= codeColIdx) continue;
        const rawCode = row[codeColIdx];
        const code = rawCode.trim();

        if (!code || seen.has(code)) continue;

        const ticker = toTicker(rawCode);
        const date = row[dateColIdx] || '';

        seen.set(code, { ticker, date, rawCode: code });
    }

    return Array.from(seen.values());
}

// ============================================
// Yahoo Finance API
// ============================================

/**
 * 指定ティッカーの終値と配当金を取得
 */
async function fetchClosingPrice(ticker, targetDateStr) {
    if (!ticker) return { price: null, dividend: null, actualDate: null, error: '無効なティッカー' };

    try {
        // 日付をUNIXタイムスタンプに変換
        // targetDateStr は "2026/02/20" 形式
        const targetDate = new Date(targetDateStr.replace(/\//g, '-'));
        const targetTs = Math.floor(targetDate.getTime() / 1000);

        // 対象日を含む前後14日間のデータを取得（配当イベントも含む）
        const startTs = targetTs - 14 * 86400;
        const endTs = targetTs + 14 * 86400;

        const apiUrl = `${YAHOO_API_BASE}${encodeURIComponent(ticker)}?period1=${startTs}&period2=${endTs}&interval=1d&events=div`;
        const proxyUrl = `${CORS_PROXY}${encodeURIComponent(apiUrl)}`;

        const response = await fetch(proxyUrl, {
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
            return { price: null, dividend: null, actualDate: null, error: `HTTP ${response.status}` };
        }

        const data = await response.json();

        if (!data.chart || !data.chart.result || data.chart.result.length === 0) {
            return { price: null, dividend: null, actualDate: null, error: 'データなし' };
        }

        const result = data.chart.result[0];
        const timestamps = result.timestamp || [];
        const closes = result.indicators?.quote?.[0]?.close || [];

        if (timestamps.length === 0 || closes.length === 0) {
            return { price: null, dividend: null, actualDate: null, error: 'チャートデータなし' };
        }

        // 対象日に最も近い日を探す（対象日以前のデータ優先）
        let bestIdx = -1;
        let bestDiff = Infinity;

        for (let i = 0; i < timestamps.length; i++) {
            if (closes[i] === null) continue;
            const diff = targetTs - timestamps[i];
            if (diff >= 0 && diff < bestDiff) {
                bestDiff = diff;
                bestIdx = i;
            }
        }

        // 対象日以前にデータがない場合、対象日以降の最も近いデータを使う
        if (bestIdx === -1) {
            for (let i = 0; i < timestamps.length; i++) {
                if (closes[i] === null) continue;
                const diff = Math.abs(targetTs - timestamps[i]);
                if (diff < bestDiff) {
                    bestDiff = diff;
                    bestIdx = i;
                }
            }
        }

        if (bestIdx === -1) {
            return { price: null, dividend: null, actualDate: null, error: '有効な終値なし' };
        }

        const actualDate = new Date(timestamps[bestIdx] * 1000);
        const formattedDate = `${actualDate.getFullYear()}/${String(actualDate.getMonth() + 1).padStart(2, '0')}/${String(actualDate.getDate()).padStart(2, '0')}`;

        // 配当金を取得（対象日に最も近い配当イベントを探す）
        let dividendAmount = null;
        const dividends = result.events?.dividends;
        if (dividends) {
            let closestDivDiff = Infinity;
            for (const [ts, divData] of Object.entries(dividends)) {
                const divTs = parseInt(ts);
                const diff = Math.abs(targetTs - divTs);
                if (diff < closestDivDiff) {
                    closestDivDiff = diff;
                    dividendAmount = divData.amount;
                }
            }
        }

        return {
            price: Math.round(closes[bestIdx] * 10) / 10,
            dividend: dividendAmount !== null ? Math.round(dividendAmount * 100) / 100 : null,
            actualDate: formattedDate,
            error: null
        };
    } catch (err) {
        return { price: null, dividend: null, actualDate: null, error: err.message };
    }
}

// ============================================
// Batch Processing
// ============================================

async function fetchAllPrices(stocks) {
    isFetching = true;
    closingPrices = {};
    errorMessages = [];

    const total = stocks.length;
    let completed = 0;

    showProgress(true);
    updateProgress(0, total, '準備中...');

    dom.fetchBtn.disabled = true;
    dom.fetchBtn.classList.add('loading');
    dom.fetchBtn.querySelector('.btn-icon').textContent = '⏳';

    for (let i = 0; i < total; i += BATCH_SIZE) {
        const batch = stocks.slice(i, i + BATCH_SIZE);

        const promises = batch.map(async (stock) => {
            const result = await fetchClosingPrice(stock.ticker, stock.date);
            closingPrices[stock.rawCode] = result;

            if (result.error) {
                errorMessages.push({
                    code: stock.rawCode,
                    ticker: stock.ticker || 'N/A',
                    error: result.error
                });
            }

            completed++;
            updateProgress(completed, total, `${stock.ticker || stock.rawCode} を取得中...`);
        });

        await Promise.all(promises);

        // レート制限対策：バッチ間に待機
        if (i + BATCH_SIZE < total) {
            await sleep(BATCH_DELAY_MS);
        }
    }

    isFetching = false;
    dom.fetchBtn.disabled = false;
    dom.fetchBtn.classList.remove('loading');
    dom.fetchBtn.querySelector('.btn-icon').textContent = '🔍';

    updateProgress(total, total, '完了');
    showResults();
    renderTable();
    dom.downloadBtn.disabled = false;
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================
// UI Updates
// ============================================

function showSection(el, show = true) {
    el.style.display = show ? '' : 'none';
    if (show) {
        el.style.animation = 'none';
        el.offsetHeight; // reflow
        el.style.animation = '';
    }
}

function showProgress(show) {
    showSection(dom.progressContainer, show);
}

function updateProgress(current, total, detail) {
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    dom.progressFill.style.width = pct + '%';
    dom.progressValue.textContent = pct + '%';
    dom.progressLabel.textContent = `${current} / ${total} 銘柄取得完了`;
    dom.progressDetail.textContent = detail;
}

function showResults() {
    const successN = Object.values(closingPrices).filter(v => v.price !== null).length;
    const errorN = errorMessages.length;
    const naCount = Object.values(closingPrices).filter(v => v.price === null).length;

    dom.successCount.textContent = successN;
    dom.naCount.textContent = naCount;
    dom.errorCount.textContent = errorN;

    showSection(dom.resultsSection);

    if (errorMessages.length > 0) {
        dom.errorLog.innerHTML = errorMessages.map(e =>
            `<div class="error-entry"><span class="error-code">${e.code}</span> (${e.ticker}) — ${e.error}</div>`
        ).join('');
        showSection(dom.errorSection);
    } else {
        showSection(dom.errorSection, false);
    }
}

function renderTable() {
    if (!parsedData) return;

    const { header, rows } = parsedData;

    // ヘッダー
    const headerLabels = [
        '基準日', '(実質上)基準日', '権利落日(普通取引)',
        '権利落日(その他取引)', '銘柄コード', '銘柄略称',
        '市場', '備考', '更新フラグ'
    ];

    const hasPrices = Object.keys(closingPrices).length > 0;
    const thLabels = [...headerLabels];
    if (hasPrices) {
        thLabels.push('終値', '配当金', '配当利回り(%)');
    }

    dom.tableHead.innerHTML = '<tr>' + thLabels.map(h => `<th>${h}</th>`).join('') + '</tr>';

    // データ行（最大200行表示）
    const displayRows = rows.slice(0, 200);
    dom.tableBody.innerHTML = displayRows.map(row => {
        const cells = row.slice(0, 9).map((cell, idx) => `<td>${escapeHTML(cell.trim())}</td>`);

        if (hasPrices) {
            const code = (row[4] || '').trim();
            const priceData = closingPrices[code];

            // 終値
            if (priceData && priceData.price !== null) {
                cells.push(`<td class="price-cell has-price">${priceData.price.toLocaleString()}</td>`);
            } else {
                cells.push(`<td class="price-cell no-price">N/A</td>`);
            }

            // 配当金
            if (priceData && priceData.dividend !== null) {
                cells.push(`<td class="price-cell has-price">${priceData.dividend.toLocaleString()}</td>`);
            } else {
                cells.push(`<td class="price-cell no-price">N/A</td>`);
            }

            // 配当利回り(%)
            if (priceData && priceData.price !== null && priceData.dividend !== null && priceData.price > 0) {
                const yieldPct = (priceData.dividend / priceData.price * 100).toFixed(2);
                cells.push(`<td class="price-cell has-price">${yieldPct}%</td>`);
            } else {
                cells.push(`<td class="price-cell no-price">N/A</td>`);
            }
        }

        return '<tr>' + cells.join('') + '</tr>';
    }).join('');

    dom.rowCount.textContent = `${rows.length} 行${rows.length > 200 ? '（200行まで表示）' : ''}`;
    showSection(dom.tableSection);
}

function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ============================================
// CSV Export
// ============================================

function generateOutputCSV() {
    if (!parsedData) return '';

    const { metadata, header, rows } = parsedData;
    const lines = [];

    // メタデータ行を保持
    for (const line of metadata) {
        lines.push(line);
    }

    // ヘッダー行（整理されたもの + 終値・配当金・配当利回り列）
    const headerLabels = [
        '基準日', '(実質上)基準日', '権利落日(普通取引)',
        '権利落日(その他の取引)', '銘柄コード', '銘柄略称',
        '市場', '備考', '更新フラグ', '終値', '配当金', '配当利回り(%)'
    ];
    lines.push(headerLabels.join(','));

    // データ行
    for (const row of rows) {
        const cells = row.slice(0, 9).map(c => {
            const val = c.trim();
            // カンマや引用符を含む場合は引用符で囲む
            if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                return '"' + val.replace(/"/g, '""') + '"';
            }
            return val;
        });

        // 終値を追加
        const code = (row[4] || '').trim();
        const priceData = closingPrices[code];
        if (priceData && priceData.price !== null) {
            cells.push(String(priceData.price));
        } else {
            cells.push('N/A');
        }

        // 配当金を追加
        if (priceData && priceData.dividend !== null) {
            cells.push(String(priceData.dividend));
        } else {
            cells.push('N/A');
        }

        // 配当利回り(%)を追加
        if (priceData && priceData.price !== null && priceData.dividend !== null && priceData.price > 0) {
            cells.push((priceData.dividend / priceData.price * 100).toFixed(2));
        } else {
            cells.push('N/A');
        }

        lines.push(cells.join(','));
    }

    return lines.join('\r\n');
}

function downloadCSV() {
    const csvContent = generateOutputCSV();
    if (!csvContent) return;

    // BOM付きUTF-8でエクスポート（Excelで文字化けしないように）
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8;' });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // ファイル名を生成
    const now = new Date();
    const timestamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    a.download = `kabuka_owarine_${timestamp}.csv`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ============================================
// File Handling
// ============================================

function handleFile(file) {
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();

    if (!['csv', 'xls', 'xlsx'].includes(ext)) {
        alert('CSV、XLS、またはXLSXファイルを選択してください。');
        return;
    }

    // Excel ファイル（XLS / XLSX）の場合
    if (ext === 'xls' || ext === 'xlsx') {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                // 最初のシートをCSVに変換
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const csvText = XLSX.utils.sheet_to_csv(firstSheet);
                processCSVText(csvText, file);
            } catch (err) {
                alert('Excelファイルの読み込みに失敗しました: ' + err.message);
                console.error(err);
            }
        };
        reader.readAsArrayBuffer(file);
        return;
    }

    // CSV ファイルの場合
    const reader = new FileReader();

    reader.onload = (e) => {
        let text = e.target.result;

        // Shift-JIS の可能性を考慮
        // FileReader の readAsText はデフォルトUTF-8
        // 文字化けしている場合は Shift-JIS で再読み込み
        if (text.includes('�') || text.includes('\ufffd')) {
            const readerSJIS = new FileReader();
            readerSJIS.onload = (e2) => {
                processCSVText(e2.target.result, file);
            };
            readerSJIS.readAsText(file, 'Shift_JIS');
            return;
        }

        processCSVText(text, file);
    };

    reader.readAsText(file, 'UTF-8');
}

function processCSVText(text, file) {
    try {
        parsedData = parseCSV(text);
        closingPrices = {};
        errorMessages = [];

        // ファイル情報を表示
        dom.fileName.textContent = file.name;
        dom.fileSize.textContent = formatFileSize(file.size);
        showSection(dom.fileInfo, true);
        dom.dropZone.style.display = 'none';

        // 銘柄数をカウント
        const codeColIdx = 4; // 銘柄コード列
        const dateColIdx = parseInt(dom.dateColumn.value);
        const stocks = getUniqueStocks(parsedData.rows, codeColIdx, dateColIdx);

        dom.stockCount.textContent = parsedData.rows.length;
        dom.uniqueStockCount.textContent = stocks.length;

        // セクション表示
        showSection(dom.settingsSection);
        showSection(dom.actionSection);
        showSection(dom.resultsSection, false);
        showSection(dom.errorSection, false);
        showSection(dom.progressContainer, false);
        dom.downloadBtn.disabled = true;

        // テーブル描画
        renderTable();

    } catch (err) {
        alert('CSV解析エラー: ' + err.message);
        console.error(err);
    }
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

function resetFile() {
    parsedData = null;
    closingPrices = {};
    errorMessages = [];

    dom.fileInput.value = '';
    showSection(dom.fileInfo, false);
    dom.dropZone.style.display = '';
    showSection(dom.settingsSection, false);
    showSection(dom.actionSection, false);
    showSection(dom.resultsSection, false);
    showSection(dom.tableSection, false);
    showSection(dom.errorSection, false);
}

// ============================================
// Event Listeners
// ============================================

// ファイルドロップ
dom.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dom.dropZone.classList.add('drag-over');
});

dom.dropZone.addEventListener('dragleave', () => {
    dom.dropZone.classList.remove('drag-over');
});

dom.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dom.dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFile(files[0]);
});

dom.dropZone.addEventListener('click', () => {
    dom.fileInput.click();
});

dom.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
});

// ファイル削除
dom.fileRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    resetFile();
});

// 日付列変更時にユニーク銘柄数を更新
dom.dateColumn.addEventListener('change', () => {
    if (!parsedData) return;
    const codeColIdx = 4;
    const dateColIdx = parseInt(dom.dateColumn.value);
    const stocks = getUniqueStocks(parsedData.rows, codeColIdx, dateColIdx);
    dom.uniqueStockCount.textContent = stocks.length;
});

// 終値取得ボタン
dom.fetchBtn.addEventListener('click', async () => {
    if (!parsedData || isFetching) return;

    const codeColIdx = 4;
    const dateColIdx = parseInt(dom.dateColumn.value);
    const stocks = getUniqueStocks(parsedData.rows, codeColIdx, dateColIdx);

    if (stocks.length === 0) {
        alert('取得対象の銘柄が見つかりません。');
        return;
    }

    await fetchAllPrices(stocks);
});

// ダウンロードボタン
dom.downloadBtn.addEventListener('click', () => {
    downloadCSV();
});
