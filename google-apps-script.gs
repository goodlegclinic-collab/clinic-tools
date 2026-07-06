/**
 * VCSS 評分工具 → Google Drive 接收器（Google Apps Script）
 * 富足診所血管醫學
 *
 * 功能：接收評分工具送來的資料 →
 *   1) 在「VCSS追蹤總表」試算表加一列（統計/研究資料庫用）
 *   2) 在病人資料夾（以病歷號搜尋子資料夾，找不到就新建）建立
 *      「病歷號_日期_VCSS.txt」，內含主訴摘要＋完整摘要
 *
 * 部署步驟見 README「Google Drive 連動」章節（5 分鐘）。
 */

const CONFIG = {
  // ① 自訂一組密碼，要跟評分工具「⚙ Google Drive 連動設定」裡填的密鑰一致
  TOKEN: '請改成你自己的密鑰',

  // ② 病人資料夾的 ID：打開該資料夾，網址 folders/ 後面那一串就是 ID。
  //    留空字串 '' = 只寫總表（放雲端硬碟根目錄）、不建病人檔案
  PARENT_FOLDER_ID: '',

  // 總表檔名（會自動建立在病人資料夾裡；已存在就直接沿用）
  MASTER_NAME: 'VCSS追蹤總表',
};

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!CONFIG.TOKEN || CONFIG.TOKEN === '請改成你自己的密鑰' || d.token !== CONFIG.TOKEN) {
      return out_({ ok: false, error: '密鑰不符（檢查腳本 CONFIG.TOKEN 與工具設定）' });
    }

    // 1) 寫入總表
    const ss = masterSheet_();
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

    // 2) 病人資料夾建檔
    let extra = '';
    if (CONFIG.PARENT_FOLDER_ID) {
      const folder = patientFolder_(String(d.chartNo));
      const name = String(d.chartNo) + '_' +
        Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmm') + '_VCSS.txt';
      folder.createFile(name, (d.cc || '') + '\n\n' + (d.summary || ''), 'text/plain');
      extra = '，病人資料夾已建檔';
    }

    return out_({ ok: true, msg: '已寫入「' + CONFIG.MASTER_NAME + '」第 ' + sh.getLastRow() + ' 列' + extra });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/** 取得（或第一次自動建立）總表 */
function masterSheet_() {
  const parent = CONFIG.PARENT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID)
    : DriveApp.getRootFolder();
  const it = parent.getFilesByName(CONFIG.MASTER_NAME);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());

  const ss = SpreadsheetApp.create(CONFIG.MASTER_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(parent);
  const sh = ss.getSheets()[0];
  sh.appendRow(['時間', '病歷號', '就診', 'VCSS總分',
    '疼痛', '靜脈曲張', '水腫', '色素沉著', '發炎', '硬化',
    '潰瘍數', '潰瘍時間', '潰瘍大小', '壓力治療',
    'CEAP', 'VDS', 'Villalta', 'CIVIQ', '主訴摘要']);
  sh.setFrozenRows(1);
  return ss;
}

/** 以病歷號搜尋病人子資料夾；找不到就在病人資料夾下新建 */
function patientFolder_(chartNo) {
  const parent = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  const found = parent.searchFolders("title contains '" + chartNo.replace(/'/g, "\\'") + "'");
  if (found.hasNext()) return found.next();
  return parent.createFolder(chartNo);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
