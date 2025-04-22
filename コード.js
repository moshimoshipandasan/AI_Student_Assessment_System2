// 環境変数（スクリプトのプロパティ）へのAPI設定
// apikeyの取得先(Ctrl+クリックで開きます→) https://aistudio.google.com/app/apikey
const apikey = PropertiesService.getScriptProperties().getProperty('apikey');
// const gemini_model = 'gemini-1.5-pro'; //高品質
const gemini_model1 = 'gemini-2.5-pro-exp-03-25'; //低速（高品質）
const gemini_model2 = 'gemini-2.5-flash-preview-04-17'; //中速（安定）
// const gemini_model2 = 'gemini-2.0-flash'; //高速（エラーが多く不安定）
// const gemini_model2 = 'gemini-1.5-flash'; //高速（安定）
const GEMINI_URL1 = `https://generativelanguage.googleapis.com/v1beta/models/${gemini_model1}:generateContent?key=${apikey}`;
const GEMINI_URL2 = `https://generativelanguage.googleapis.com/v1beta/models/${gemini_model2}:generateContent?key=${apikey}`;

// プロンプト
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_PROMPT = SS.getSheetByName('プロンプト');
const constraints = SHEET_PROMPT.getRange(5, 2).getValue();
const SYSTEM_PROMPT = `${constraints}`;

function generateGeminiPrompt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = [sheet.getRange('B2').getValue(), sheet.getRange('B3').getValue(), sheet.getRange('B4').getValue()];
const baseText = `
# タスクの説明
${values.join('、')}

## 役割・目標
あなたは「Gemini 用プロンプト自動生成アシスタント」です。  
ユーザーが指定したタスクに対し、Gemini が理解・実行しやすい **日本語プロンプト** を作成してください。

## 視点・対象
- 主対象:Gemini AI モデルを利用するユーザー  
- 副対象:Gemini AI モデル本体（プロンプトの受け手）

## 制約条件
1. Gemini のガイドラインとトークン上限を守る  
2. 不要な解説や英語混在は避け、**日本語で簡潔**に  
3. 出力は**下記フォーマットを厳守**し、プレースホルダは <> で示す  
4. 法律・倫理に反する内容を含めない  
5. Markdown 見出しの「#」「##」は必ず半角

## 処理手順 (Chain of Thought 略述)
1. ユーザー入力から目的・条件を抽出  
2. 役割・目標／視点・対象／制約条件を整理  
3. 必要入力・期待出力を決定  
4. フォーマットに沿ってプロンプト本文を生成  
5. 生成内容を再点検し、冗長表現を削除

## 出力フォーマット（この形を厳守）
# <タスク名> Gemini Prompt

## 役割・目標
<ここに記述>

## 視点・対象
<ここに記述>

## 制約条件
1. <条件>
2. <条件>
   …

## 処理手順 (Chain of Thought)
1. <手順>
2. <手順>
   …

## 入力文
<必要な入力情報>

## 出力文
<期待される出力形式>

### ユーザー入力
<タスク説明を引用してプロンプトを生成してください>
`;

  const payload = {
    'contents': [{
      'parts': [{
        'text': baseText
      }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(GEMINI_URL1, options);
    const responseJson = JSON.parse(response.getContentText());
    
    if (responseJson && responseJson.candidates && responseJson.candidates.length > 0) {
      const generatedPrompt = responseJson.candidates[0].content.parts[0].text;
      sheet.getRange('B5').setValue(generatedPrompt);
      return generatedPrompt;
    } else {
      const errorMessage = 'No response from Gemini API';
      sheet.getRange('B5').setValue(errorMessage);
      return errorMessage;
    }
  } catch (e) {
    const errorMessage = 'Error retrieving response: ' + e.toString();
    sheet.getRange('B5').setValue(errorMessage);
    return errorMessage;
  }
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('添削メニュー')
    .addItem('1.データの読込み', 'checkDriveChanges')
    .addItem('2.読み込みデータ整理', 'cleanupDuplicateData')
    .addItem('3.評価・添削結果をメール送信', 'sendEmailsWithGeminiAndSubmission')
    .addSeparator()
    .addItem('0.プロンプト自動生成', 'generateGeminiPrompt')
    .addToUi();
}

function aiscore(values) {
  if (!values || (Array.isArray(values) && values.length === 0)) {
    return null; // 値が空の場合は何も返さない
  }

  if (!Array.isArray(values)) {
    values = [[values]];
  }
  const flatValues = values.flat();
  const prompt = flatValues.join(',');

  const payload = {
    systemInstruction: {
      role: "model",
      parts:[{
        text: SYSTEM_PROMPT
      }]
    },
    contents: [{
      role: "user",
      parts:[{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.0,
      top_p: 1,        // または省略
      // top_k: 40,     // 省略可。両方指定するなら残す
      max_output_tokens: 2048
    }
    // generationConfig: {
    // https://ai.google.dev/api/python/google/generativeai/types/GenerationConfig
      // temperature: 0.1, // 生成するテキストのランダム性を制御
      // top_p: 0.1, // 生成に使用するトークンの累積確率を制御
      // top_k: 40, // 生成に使用するトップkトークンを制御
      // max_output_tokens: 8192 // 最大出力トークン数を指定
    // }
  },
  options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(GEMINI_URL2, options),
          responseJson = JSON.parse(response.getContentText());

    if (responseJson && responseJson.candidates && responseJson.candidates.length > 0) {
      return responseJson.candidates[0].content.parts[0].text;
    } else {
      return 'No response from Gemini API';
    }
  } catch (e) {
    return 'Error retrieving response: ' + e.toString();
  }
}

/**
 * 評価・添削シートで同じFile Nameを持つ行のうち、Contentsが空の行を削除する関数
 * @return {number} 削除された行数
 */
function cleanupDuplicateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('評価・添削');
  
  if (!sheet) {
    Logger.log('評価・添削シートが見つかりません。');
    Browser.msgBox('エラー', '評価・添削シートが見つかりません。', Browser.Buttons.OK);
    return 0;
  }
  
  // データを取得
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 各列のインデックスを取得
  const fileNameIndex = headers.indexOf('File Name');
  const contentsIndex = headers.indexOf('Contents');
  
  if (fileNameIndex === -1 || contentsIndex === -1) {
    Logger.log('必要なカラム（File Name または Contents）が見つかりません。');
    Browser.msgBox('エラー', '必要なカラム（File Name または Contents）が見つかりません。', Browser.Buttons.OK);
    return 0;
  }
  
  // File Name ごとにデータをグループ化
  const fileNameGroups = {};
  for (let i = 1; i < data.length; i++) {
    const fileName = data[i][fileNameIndex];
    if (fileName) {
      if (!fileNameGroups[fileName]) {
        fileNameGroups[fileName] = [];
      }
      fileNameGroups[fileName].push({
        rowIndex: i + 1, // スプレッドシートの行番号（1始まり）
        contents: data[i][contentsIndex],
        rowData: data[i]
      });
    }
  }
  
  // 削除する行を特定（降順にソートして後ろから削除）
  const rowsToDelete = [];
  for (const fileName in fileNameGroups) {
    const rows = fileNameGroups[fileName];
    if (rows.length > 1) {
      // 同じFile Nameを持つ行が複数ある場合
      // Contentsが空の行を特定
      const emptyContentRows = rows.filter(row => !row.contents);
      // Contentsが空でない行が少なくとも1つある場合のみ、空の行を削除対象に追加
      if (rows.length - emptyContentRows.length > 0) {
        emptyContentRows.forEach(row => {
          rowsToDelete.push(row.rowIndex);
        });
      }
    }
  }
  
  // 行を降順にソート（後ろから削除するため）
  rowsToDelete.sort((a, b) => b - a);
  
  // 行を削除
  rowsToDelete.forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });
  
  const message = `データ整理が完了しました。${rowsToDelete.length}行の重複データを削除しました。`;
  Logger.log(message);
  Browser.msgBox('処理完了', message, Browser.Buttons.OK);
  
  return rowsToDelete.length;
}