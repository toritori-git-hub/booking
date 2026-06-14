// ═══════════════════════════════════════════════
//  クリーンハンター 仮予約受付 GAS doPost
//  Googleスプレッドシートに予約データを保存する
// ═══════════════════════════════════════════════

// ▼ ここにスプレッドシートIDを入れる（URLの /d/〇〇〇/edit の〇〇〇部分）
const SHEET_ID   = '1MnC06qXeweHsuX9wbnYRvpFzOqenL8IkqBObOyZ8n-I';
const SHEET_NAME = '予約一覧';

// ▼ 通知メールの送信先（担当者メールアドレス）
const NOTIFY_EMAIL = 'info@clean-hunter.com';

// 見出し行の定義（列順）
const HEADERS = [
  '送信日時',
  'LINE名またはお名前',
  '希望連絡方法',
  'メールアドレス',
  '電話番号',
  '作業場所',
  '希望作業',
  'エアコン台数',
  'お掃除機能',
  'エアコンメーカー型番',
  'ドラム式メーカー型番',
  '洗濯機上ラック有無',
  '設置写真送付状況',
  'お困りごと',
  'お急ぎ度',
  '初回利用',
  '第1希望日',
  '第1希望時間帯',
  '第2希望日',
  '第2希望時間帯',
  '駐車場',
  '備考'
];

function doPost(e) {
  try {
    // ── リクエストボディをJSONとして解析 ──
    var raw = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    // ── スプレッドシートを取得 ──
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // ── 1行目が空なら見出しを挿入 ──
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      // 見出し行を太字・背景色で装飾
      var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1a73e8');
      headerRange.setFontColor('#ffffff');
    }

    // ── 保存するデータ行を作成（HEADERS順に並べる） ──
    var row = [
      data.submitted_at     || new Date().toLocaleString('ja-JP'),
      data.line_name        || '',
      data.contact_method   || '',
      data.email            || '',
      data.phone            || '',
      data.address          || '',
      data.menu             || '',
      data.aircon_count     || '',
      data.aircon_clean_func || '',
      data.aircon_info      || '',
      data.washer_info      || '',
      data.washer_rack      || '',
      data.washer_photo     || '',
      data.trouble          || '',
      data.urgency          || '',
      data.first_visit      || '',
      data.first_date       || '',
      data.first_time       || '',
      data.second_date      || '',
      data.second_time      || '',
      data.parking          || '',
      data.notes            || ''
    ];

    sheet.appendRow(row);

    // ── 担当者へ通知メールを送信 ──
    var subject = '【仮予約】' + (data.line_name || '不明') + ' 様より仮予約が届きました';
    var body = [
      '■ 仮予約が届きました',
      '',
      '送信日時　　　：' + (data.submitted_at || ''),
      'お名前　　　　：' + (data.line_name || ''),
      '連絡方法　　　：' + (data.contact_method || ''),
      'メールアドレス：' + (data.email || '（なし）'),
      '電話番号　　　：' + (data.phone || '（なし）'),
      '作業場所　　　：' + (data.address || ''),
      '希望作業　　　：' + (data.menu || ''),
      'エアコン台数　：' + (data.aircon_count || ''),
      'お掃除機能　　：' + (data.aircon_clean_func || ''),
      'エアコン型番　：' + (data.aircon_info || ''),
      'ドラム式型番　：' + (data.washer_info || ''),
      'ラック有無　　：' + (data.washer_rack || ''),
      '写真送付状況　：' + (data.washer_photo || ''),
      'お困りごと　　：' + (data.trouble || ''),
      'お急ぎ度　　　：' + (data.urgency || ''),
      '初回利用　　　：' + (data.first_visit || ''),
      '第1希望日　　：' + (data.first_date || '') + ' ' + (data.first_time || ''),
      '第2希望日　　：' + (data.second_date || '') + ' ' + (data.second_time || ''),
      '駐車場　　　　：' + (data.parking || ''),
      '備考　　　　　：' + (data.notes || ''),
      '',
      '─────────────────',
      '※ スプレッドシートにも自動保存されています。',
      '※ このメールはフォーム送信時に自動送信されました。'
    ].join('\n');

    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);

    // ── 成功レスポンスを返す ──
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // ── エラーレスポンスを返す ──
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── テスト用（GASエディタから手動実行して動作確認） ──
function testDoPost() {
  var testData = {
    submitted_at:      '2026/06/13 12:00:00',
    line_name:         'テスト太郎',
    contact_method:    'LINE',
    email:             '',
    phone:             '',
    address:           '神戸市東灘区深江南町',
    menu:              'エアコンクリーニング',
    aircon_count:      '2台',
    aircon_clean_func: 'あり',
    aircon_info:       'ダイキン F-40XEXV',
    washer_info:       '',
    washer_rack:       '',
    washer_photo:      '',
    trouble:           'エアコンがカビ臭い',
    urgency:           '1週間以内希望',
    first_visit:       '初めて',
    first_date:        '2026-06-20',
    first_time:        '午前',
    second_date:       '2026-06-21',
    second_time:       'どちらでも可',
    parking:           'あり',
    notes:             ''
  };

  var mockEvent = {
    postData: { contents: JSON.stringify(testData) }
  };

  var result = doPost(mockEvent);
  Logger.log(result.getContent());
}
