/**
 * 富足診所 臨床工具 → Google Drive 接收器（Google Apps Script）
 * 服務對象：VCSS 評分工具 ＋ 門診對話分析器
 *
 * 動作（依 payload 的 action 欄位分流；舊版工具沒帶 action = 'save'，完全相容）：
 *   save         VCSS 評分上傳 → 「VCSS追蹤總表」加一列 ＋ 病人資料夾建立 VCSS.txt
 *   saveAnalysis 門診對話分析上傳 → 「門診對話分析總表」加一列 ＋ 病人資料夾建立 門診分析.txt
 *   list         以病歷號找病人資料夾 → 回傳文字稿／紀錄檔清單（含日期子資料夾一層）
 *   get          以檔案 ID 取得文字內容（.txt/.md/Google 文件）
 *   latest       回傳資料夾內「最新的文字稿」；帶 since 時間戳可做自動監看
 *   outcome      事後回填成交結果 → 更新分析總表該病歷號最近一筆
 *   history      回傳分析總表最近數百筆（給工具端算成交率與品質×成交洞察）
 *
 * 部署步驟見 README「Google Drive 連動」章節。
 * ※ 更新這支腳本後，記得「部署 → 管理部署作業 → 編輯 → 新版本」才會生效；
 *   第一次用到讀取 Google 文件的功能會再要求一次授權。
 */

const CONFIG = {
  // ① 自訂一組密碼，要跟工具「⚙ Google Drive 連動設定」裡填的密鑰一致
  TOKEN: '請改成你自己的密鑰',

  // ② 病人資料夾的 ID：打開該資料夾，網址 folders/ 後面那一串就是 ID。
  //    留空字串 '' = 只寫總表（放雲端硬碟根目錄）、不建病人檔案；list/latest 會全 Drive 搜尋
  PARENT_FOLDER_ID: '',

  // 總表檔名（會自動建立；已存在就直接沿用）
  MASTER_NAME: 'VCSS追蹤總表',
  ANALYSIS_MASTER_NAME: '門診對話分析總表',

  // 病歷號在資料夾名稱裡的補零位數（0001550 ← 1550）
  CHART_PAD: 7,
};

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!CONFIG.TOKEN || CONFIG.TOKEN === '請改成你自己的密鑰' || d.token !== CONFIG.TOKEN) {
      return out_({ ok: false, error: '密鑰不符（檢查腳本 CONFIG.TOKEN 與工具設定）' });
    }
    switch (d.action || 'save') {
      case 'save':         return saveVcss_(d);
      case 'saveAnalysis': return saveAnalysis_(d);
      case 'list':         return listFiles_(d);
      case 'get':          return getFile_(d);
      case 'latest':       return latest_(d);
      case 'outcome':      return outcome_(d);
      case 'history':      return history_(d);
      default:             return out_({ ok: false, error: '未知動作：' + d.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ================= VCSS 評分上傳（原有行為，未變） ================= */

function saveVcss_(d) {
  const ss = masterSheet_(CONFIG.MASTER_NAME, ['時間', '病歷號', '就診', 'VCSS總分',
    '疼痛', '靜脈曲張', '水腫', '色素沉著', '發炎', '硬化',
    '潰瘍數', '潰瘍時間', '潰瘍大小', '壓力治療',
    'CEAP', 'VDS', 'Villalta', 'CIVIQ', '主訴摘要']);
  const sh = ss.getSheets()[0];
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  const dm = d.domains || {};
  sh.appendRow([
    ts, String(d.chartNo || ''), d.visitType || '',
    d.vcss,
    dm.pain, dm.veins, dm.edema, dm.pigment, dm.inflam,
    dm.indur, dm.ulcerN, dm.ulcerT, dm.ulcerS, dm.comp,
    d.ceap || '', d.vds, d.villalta, d.civiq,
    d.cc || ''
  ]);

  let extra = '';
  if (CONFIG.PARENT_FOLDER_ID) {
    const folder = patientFolder_(String(d.chartNo));
    const name = String(d.chartNo) + '_' +
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmm') + '_VCSS.txt';
    folder.createFile(name, (d.cc || '') + '\n\n' + (d.summary || ''), 'text/plain');
    extra = '，病人資料夾已建檔';
  }
  return out_({ ok: true, msg: '已寫入「' + CONFIG.MASTER_NAME + '」第 ' + sh.getLastRow() + ' 列' + extra });
}

/* ================= 門診對話分析上傳／成交回填／統計 ================= */

const ANALYSIS_HEADERS = ['時間', '病歷號', '就診', 'VCSS總分', 'CEAP', 'VDS',
  '主訴摘要', '治療討論', '病人疑慮', '醫囑待辦', '決策狀態',
  '品質總分', 'SPIN', '異議處理', '下一步', '成交狀態', '成交項目', '成交金額'];

function saveAnalysis_(d) {
  const sh = masterSheet_(CONFIG.ANALYSIS_MASTER_NAME, ANALYSIS_HEADERS).getSheets()[0];
  const headers = ensureHeaders_(sh, ANALYSIS_HEADERS);
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  appendByHeaders_(sh, headers, {
    '時間': ts, '病歷號': String(d.chartNo || ''), '就診': d.visitType || '',
    'VCSS總分': d.vcss, 'CEAP': d.ceap || '', 'VDS': d.vds,
    '主訴摘要': d.cc || '', '治療討論': d.treatments || '', '病人疑慮': d.concerns || '',
    '醫囑待辦': d.orders || '', '決策狀態': d.decision || '',
    '品質總分': d.quality, 'SPIN': d.spin || '', '異議處理': d.objection || '',
    '下一步': d.cta || '', '成交狀態': d.outcome || '追蹤中',
    '成交項目': d.outcomeItem || '', '成交金額': d.outcomeAmount || ''
  });

  let extra = '';
  if (CONFIG.PARENT_FOLDER_ID) {
    const folder = patientFolder_(String(d.chartNo));
    const name = String(d.chartNo) + '_' +
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmm') + '_門診分析.txt';
    folder.createFile(name, (d.cc || '') + '\n\n' + (d.summary || ''), 'text/plain');
    extra = '，病人資料夾已建檔';
  }
  return out_({ ok: true, msg: '已寫入「' + CONFIG.ANALYSIS_MASTER_NAME + '」第 ' + sh.getLastRow() + ' 列' + extra });
}

/** 事後回填成交：更新該病歷號「最近一筆」的成交欄位；找不到就補一列 */
function outcome_(d) {
  const sh = masterSheet_(CONFIG.ANALYSIS_MASTER_NAME, ANALYSIS_HEADERS).getSheets()[0];
  const headers = ensureHeaders_(sh, ANALYSIS_HEADERS);
  const chart = String(d.chartNo || '');
  if (!chart) return out_({ ok: false, error: '缺少病歷號' });
  const colChart = headers.indexOf('病歷號') + 1;
  const set = { '成交狀態': d.outcome || '', '成交項目': d.outcomeItem || '', '成交金額': d.outcomeAmount || '' };
  const last = sh.getLastRow();
  if (last >= 2) {
    const vals = sh.getRange(2, colChart, last - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === chart) {
        const row = i + 2;
        Object.keys(set).forEach(function (h) {
          const c = headers.indexOf(h) + 1;
          // 狀態一定更新；項目／金額留空就不覆蓋舊值
          if (c > 0 && (set[h] !== '' || h === '成交狀態')) sh.getRange(row, c).setValue(set[h]);
        });
        return out_({ ok: true, msg: '已更新病歷號 ' + chart + ' 第 ' + row + ' 列的成交結果' });
      }
    }
  }
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  appendByHeaders_(sh, headers, Object.assign({ '時間': ts, '病歷號': chart }, set));
  return out_({ ok: true, msg: '總表沒有這個病歷號的分析紀錄，已補一列成交紀錄' });
}

/** 回傳最近數百筆分析紀錄（工具端計算成交率用） */
function history_(d) {
  const sh = masterSheet_(CONFIG.ANALYSIS_MASTER_NAME, ANALYSIS_HEADERS).getSheets()[0];
  const headers = ensureHeaders_(sh, ANALYSIS_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return out_({ ok: true, rows: [] });
  const n = Math.min(last - 1, 300);
  const data = sh.getRange(last - n + 1, 1, n, headers.length).getValues();
  const col = function (h) { return headers.indexOf(h); };
  const rows = data.map(function (r) {
    return {
      date: String(r[col('時間')] || ''), chartNo: String(r[col('病歷號')] || ''),
      visitType: String(r[col('就診')] || ''), decision: String(r[col('決策狀態')] || ''),
      quality: r[col('品質總分')], spin: String(r[col('SPIN')] || ''),
      objection: String(r[col('異議處理')] || ''), cta: String(r[col('下一步')] || ''),
      outcome: String(r[col('成交狀態')] || ''), outcomeAmount: String(r[col('成交金額')] || '')
    };
  });
  return out_({ ok: true, rows: rows });
}

/** 表頭防呆：舊版總表缺的欄位自動補在最後面，回傳目前表頭 */
function ensureHeaders_(sh, wanted) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let next = cur.filter(function (h) { return h !== ''; });
  wanted.forEach(function (h) {
    if (next.indexOf(h) < 0) {
      sh.getRange(1, next.length + 1).setValue(h);
      next.push(h);
    }
  });
  return next;
}
function appendByHeaders_(sh, headers, obj) {
  sh.appendRow(headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

/* ================= 文字稿：列出／取得／最新 ================= */

/** 病人資料夾＋一層子資料夾（日期夾）內的文字類檔案清單，新→舊 */
function listFiles_(d) {
  const folder = findPatientFolder_(String(d.chartNo || ''));
  if (!folder) return out_({ ok: false, error: '找不到病歷號「' + d.chartNo + '」的資料夾' });
  const files = collectTextFiles_(folder).sort(function (a, b) {
    return b.updatedMs - a.updatedMs;
  }).slice(0, 25).map(function (f) {
    return { id: f.id, name: f.name, type: f.type, path: f.path, updated: f.updated };
  });
  return out_({ ok: true, folder: folder.getName(), folderId: folder.getId(), files: files });
}

/** 取得單一檔案文字內容（txt/md/Google 文件） */
function getFile_(d) {
  const file = DriveApp.getFileById(String(d.fileId));
  const content = readTextContent_(file);
  if (content === null) return out_({ ok: false, error: '不支援的檔案格式：' + file.getMimeType() + '（請用 .txt 或 Google 文件）' });
  return out_({
    ok: true, name: file.getName(), content: content,
    updated: file.getLastUpdated().toISOString()
  });
}

/** 最新的文字稿；帶 since（ISO 時間）時，沒有更新就回 noNew，供前端每 30 秒輪詢 */
function latest_(d) {
  const folder = findPatientFolder_(String(d.chartNo || ''));
  if (!folder) return out_({ ok: false, error: '找不到病歷號「' + d.chartNo + '」的資料夾' });
  const files = collectTextFiles_(folder)
    .filter(function (f) { return f.type === 'transcript'; })
    .sort(function (a, b) { return b.updatedMs - a.updatedMs; });
  if (!files.length) return out_({ ok: true, noNew: true, msg: '資料夾內沒有文字稿' });
  const newest = files[0];
  if (d.since && newest.updatedMs <= new Date(d.since).getTime()) {
    return out_({ ok: true, noNew: true });
  }
  const file = DriveApp.getFileById(newest.id);
  const content = readTextContent_(file);
  if (content === null) return out_({ ok: true, noNew: true });
  return out_({ ok: true, name: newest.name, content: content, updated: newest.updated });
}

/* ================= 共用 ================= */

/** 以病歷號找病人資料夾：資料夾名完全等於病歷號（含補零）優先，其次名稱包含 */
function findPatientFolder_(chartNo) {
  if (!chartNo) return null;
  const padded = ('0000000000' + chartNo).slice(-CONFIG.CHART_PAD);
  const exact = [chartNo, padded];
  const it = DriveApp.searchFolders("title contains '" + chartNo.replace(/'/g, "\\'") + "'");
  let fallback = null;
  while (it.hasNext()) {
    const f = it.next();
    if (exact.indexOf(f.getName()) >= 0) return f;
    if (!fallback && f.getName().indexOf(chartNo) >= 0) fallback = f;
  }
  return fallback;
}

/** 收集資料夾（含一層子資料夾）內的文字類檔案 */
function collectTextFiles_(folder) {
  const out = [];
  scanFolder_(folder, '', out);
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    scanFolder_(sub, sub.getName(), out);
  }
  return out;
}
function scanFolder_(folder, path, out) {
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mime = f.getMimeType();
    const name = f.getName();
    const isText = mime === 'text/plain' || mime === 'text/markdown' ||
      mime === 'application/vnd.google-apps.document' ||
      /\.(txt|md)$/i.test(name);
    if (!isText) continue;
    out.push({
      id: f.getId(), name: name, path: path,
      type: /VCSS/i.test(name) ? 'vcss' : /門診分析|分析/.test(name) ? 'analysis' : 'transcript',
      updated: f.getLastUpdated().toISOString(),
      updatedMs: f.getLastUpdated().getTime()
    });
  }
}

/** 讀出文字內容；不支援的格式回 null */
function readTextContent_(file) {
  const mime = file.getMimeType();
  if (mime === 'application/vnd.google-apps.document') {
    return DocumentApp.openById(file.getId()).getBody().getText();
  }
  if (mime === 'text/plain' || mime === 'text/markdown' || /\.(txt|md)$/i.test(file.getName())) {
    return file.getBlob().getDataAsString('UTF-8');
  }
  return null;
}

/** 取得（或第一次自動建立）總表 */
function masterSheet_(name, headers) {
  const parent = CONFIG.PARENT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID)
    : DriveApp.getRootFolder();
  const it = parent.getFilesByName(name);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());

  const ss = SpreadsheetApp.create(name);
  DriveApp.getFileById(ss.getId()).moveTo(parent);
  const sh = ss.getSheets()[0];
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  return ss;
}

/** 以病歷號搜尋病人子資料夾；找不到就在病人資料夾下新建 */
function patientFolder_(chartNo) {
  const found = findPatientFolder_(chartNo);
  if (found) return found;
  const parent = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  return parent.createFolder(chartNo);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
