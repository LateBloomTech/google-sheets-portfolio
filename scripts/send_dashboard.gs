// ===============================================================
// === Googleスプレッドシートの特定シートをPDF化してメール送信する関数 ===
// ===============================================================
function sendDashboardAsPDF() {
  
  // --- ▼ ここから設定項目 ▼ ---

  const sheetName = "Dashboard"; // PDF化したいシートの名前を指定
  const recipientEmail = PropertiesService.getScriptProperties().getProperty('RECIPIENT_EMAIL');//スクリプト プロパティ → RECIPIENT_EMAIL : your.name@example.com
  if (!recipientEmail) throw new Error('Script property RECIPIENT_EMAIL is not set');
  const emailSubject = "【日次レポート】ダッシュボード"; // メールの件名
  const emailBody = "本日分のダッシュボードを添付します。\n\nご確認ください。"; // メールの本文

  // --- ▲ ここまで設定項目 ▲ ---

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      console.error("エラー: 指定されたシートが見つかりません。シート名を確認してください: " + sheetName);
      return; // シートが見つからなければ処理を終了
    }

    const sheetId = sheet.getSheetId(); // シートのIDを取得

    // PDFをエクスポートするためのURLを作成
    const url = spreadsheet.getUrl().replace(/edit$/, '') +
      'export?exportFormat=pdf&format=pdf' +
      '&size=A4' + // 用紙サイズ (A4, B5, letterなど)
      '&portrait=true' + // true=縦向き, false=横向き
      '&fitw=true' + // ページの幅に合わせる
      '&sheetnames=false&printtitle=false' + // シート名やタイトルを非表示
      '&gridlines=false' + // グリッド線を非表示
      '&gid=' + sheetId; // シートIDを指定

    const token = ScriptApp.getOAuthToken();
    const options = {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    };

    // URLからPDFデータを取得
    const response = UrlFetchApp.fetch(url, options);
    const pdfBlob = response.getBlob().setName(spreadsheet.getName() + "（" + sheetName + "）.pdf");

    // メールを送信
    GmailApp.sendEmail(recipientEmail, emailSubject, emailBody, {
      attachments: [pdfBlob]
    });
    
    console.log("PDFのメール送信が正常に完了しました。");

  } catch (e) {
    // エラーが発生した場合の処理
    console.error("エラーが発生しました: " + e.message);
    // エラー内容を自分宛にメールで通知するなどの処理も可能
    GmailApp.sendEmail(recipientEmail, "【エラー通知】ダッシュボード自動送信スクリプト", "スクリプトの実行中にエラーが発生しました。\n\nエラー内容:\n" + e.message);
  }
}