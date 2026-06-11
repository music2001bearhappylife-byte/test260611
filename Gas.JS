function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  const labels = [
    'とても暑い（汗ばむほど暑い）',
    'やや暑い（少し暑さを感じる）',
    'ちょうどよい（快適）',
    'やや寒い（少し肌寒い）',
    'とても寒い（震えてしまうほど寒い）'
  ];
  const counts = {};
  labels.forEach(l => counts[l] = 0);

  // 今日の日付を取得（日本時間）
  const now   = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 1行目はヘッダーなので2行目から集計
  // A列：タイムスタンプ、B列：回答
  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0];
    const answer    = String(data[i][1]).trim();

    // タイムスタンプを日本時間の日付に変換
    if (!timestamp) continue;
    const rowDate = Utilities.formatDate(new Date(timestamp), 'Asia/Tokyo', 'yyyy/MM/dd');

    // 今日の回答だけ集計
    if (rowDate !== today) continue;
    if (counts[answer] !== undefined) counts[answer]++;
  }

  const total = Object.values(counts).reduce((a, b) => a + b, 0);

  const result = {
    total:      total,
    counts:     counts,
    updated_at: now.toISOString()
  };

  const callback = e && e.parameter && e.parameter.callback;
  const json     = JSON.stringify(result);
  const output   = callback ? callback + '(' + json + ')' : json;
  const mime     = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(output).setMimeType(mime);
}
