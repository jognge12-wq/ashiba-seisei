/**
 * 足場 是正 集計ツール — データ保存用バックエンド（Google Apps Script）
 * ====================================================================
 * 【セットアップ手順】
 *  1. Google スプレッドシートを新規作成する（名前は何でもOK。例：足場是正データ）
 *  2. メニュー「拡張機能 → Apps Script」を開く
 *  3. エディタの中身を全部消し、このファイルの内容を貼り付けて保存
 *  4. 右上「デプロイ → 新しいデプロイ」→ 種類「ウェブアプリ」
 *       - 次のユーザーとして実行：自分
 *       - アクセスできるユーザー：全員
 *  5. 「デプロイ」→ 表示される「ウェブアプリのURL」をコピー
 *  6. その URL を伝えてもらえれば、ツール側に組み込みます
 *
 *  ※ データはこのスプレッドシート内（tallies / meta シート）に保存されます。
 *  ※ tallies・meta シートは初回アクセス時に自動で作られます。
 *  ※ このコードに鍵・パスワードの類は含まれません。
 */

var TALLY_SHEET = 'tallies';
var META_SHEET = 'meta';
var TALLY_HEADERS = ['id', 'date', 'genba', 'koumuten', 'tantou', 'gyosha', 'memo', 'updatedAt', 'counts'];

function doGet(e) {
  return jsonOut(readState());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var state = JSON.parse(e.postData.contents);
    writeState(state);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function readState() {
  var tsh = getSheet(TALLY_SHEET);
  var msh = getSheet(META_SHEET);
  var tallies = [];
  var lastRow = tsh.getLastRow();
  if (lastRow >= 2) {
    var values = tsh.getRange(2, 1, lastRow - 1, TALLY_HEADERS.length).getValues();
    values.forEach(function (r) {
      if (!r[0]) return;
      var counts = {};
      try { counts = JSON.parse(r[8] || '{}'); } catch (e) {}
      tallies.push({
        id: String(r[0]), date: String(r[1]), genba: String(r[2]),
        koumuten: String(r[3]), tantou: String(r[4]), gyosha: String(r[5]),
        memo: String(r[6]), updatedAt: String(r[7]), counts: counts
      });
    });
  }
  var masters = null;
  try {
    var raw = msh.getRange('B2').getValue();
    masters = raw ? JSON.parse(raw) : null;
  } catch (e) {}
  return {
    exists: (lastRow >= 2) || !!masters,
    branch: msh.getRange('B1').getValue() || '',
    masters: masters,
    tallies: tallies,
    lastExport: msh.getRange('B3').getValue() || null
  };
}

function writeState(state) {
  var tsh = getSheet(TALLY_SHEET);
  var msh = getSheet(META_SHEET);
  tsh.clear();
  tsh.getRange(1, 1, 1, TALLY_HEADERS.length).setValues([TALLY_HEADERS]);
  var tallies = (state && state.tallies) || [];
  if (tallies.length) {
    var rows = tallies.map(function (t) {
      return [t.id || '', t.date || '', t.genba || '', t.koumuten || '', t.tantou || '',
        t.gyosha || '', t.memo || '', t.updatedAt || '', JSON.stringify(t.counts || {})];
    });
    tsh.getRange(2, 1, rows.length, TALLY_HEADERS.length).setValues(rows);
  }
  msh.getRange('A1').setValue('branch');
  msh.getRange('B1').setValue((state && state.branch) || '');
  msh.getRange('A2').setValue('masters');
  msh.getRange('B2').setValue(JSON.stringify((state && state.masters) || {}));
  msh.getRange('A3').setValue('lastExport');
  msh.getRange('B3').setValue((state && state.lastExport) || '');
}
