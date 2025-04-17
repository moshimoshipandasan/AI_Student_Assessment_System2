function sendEmailsWithGeminiAndSubmission() {
  // スプレッドシートを開く
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("評価・添削");
  
  // データを取得
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 各列のインデックスを取得
  const idIndex = headers.indexOf("ID");
  const fileNameIndex = headers.indexOf("File Name");
  const fileUrlIndex = headers.indexOf("File Url");
  const geminiIndex = headers.indexOf("aiscore関数");
  const emailIndex = headers.indexOf("メールアドレス");
  const changeDateIndex = headers.indexOf("Change Date");
  const sendStatusIndex = 9; // J列は0から数えて9番目
  
  // 現在の日付を取得
  const date = new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' });
  
  // 各行に対して処理を行う
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[emailIndex];
    const geminiContent = row[geminiIndex];
    const fileUrl = row[fileUrlIndex];
    const fileName = row[fileNameIndex];
    const id = row[idIndex];
    const changeDate = row[changeDateIndex];
    const sendStatus = row[sendStatusIndex];

    // J列のセルが「送信済」でない場合のみ処理を続行
    if (sendStatus !== "送信済") {
      // メールの件名を作成（日付を含む）
      const subject = `レポート評価・添削結果: ${date}`;
      let body;
      
      if (!geminiContent) {
        // aiscore関数セルの値が空白の場合
        body = `
今回のレポート提出はありませんでした。

ご確認ください。
よろしくお願いいたします。
`;
      } else {
        // 通常のメール内容
        body = `
提出物に対する 評価・添削の結果です。

提出物（${changeDate}）のURL：
${fileUrl || "URLなし"}

評価・添削の結果:
${geminiContent}

ご質問がありましたら、お気軽にお問い合わせください。
よろしくお願いいたします。
`;
      }
      
      let newSendStatus = "送信できませんでした";
      
      // メールを送信
      if (email) {
        try {
          MailApp.sendEmail(email, subject, body);
          newSendStatus = "送信済";
          Logger.log(`メールを送信しました: ${email}`);
        } catch (error) {
          Logger.log(`メール送信エラー: ${error}`);
        }
      } else {
        Logger.log(`メールの送信をスキップしました（メールアドレスなし）: ID ${id}`);
      }
      
      // J列（10列目）に送信状況を記録
      sheet.getRange(i + 1, 10).setValue(newSendStatus);
    } else {
      Logger.log(`メールの送信をスキップしました（既に送信済）: ID ${id}`);
    }
  }
}

// トリガーを設定する関数（必要に応じて使用）
function setDailyTrigger() {
  ScriptApp.newTrigger('sendEmailsWithGeminiAndSubmission')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}