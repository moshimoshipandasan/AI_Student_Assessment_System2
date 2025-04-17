function extractFolderID() {
  // "フォルダID"シートを取得
  var SHEET_FOLDERID = SS.getSheetByName("フォルダURL");
  // A1セルの値を取得
  var cellValue = SHEET_FOLDERID.getRange("A1").getValue();  
  // 正規表現を使用してフォルダIDを抽出
  var regex = /https:\/\/drive\.google\.com\/drive\/folders\/([a-zA-Z0-9_-]+)/;
  var match = cellValue.match(regex);
  // マッチした場合、フォルダIDを返す。マッチしない場合はnullを返す
  return match ? match[1] : null;
}

// トリガーの設定を行う関数
function setupTrigger() {
  ScriptApp.newTrigger('checkDriveChanges')
    .timeBased()
    .everyDays(1) // 1日ごとにトリガーを実行
    .create();
}

function checkDriveChanges() {
  const folderId = extractFolderID();
  if (!folderId) {
    Logger.log('フォルダIDが見つかりません。');
    return;
  }
  
  const folder = DriveApp.getFolderById(folderId);
  if (!folder) {
    Logger.log('指定されたIDのフォルダが見つかりません。');
    return;
  }
  
  const files = folder.getFiles();
  const sheet = SS.getSheetByName('評価・添削');

  // ヘッダー行を追加または確認
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'File Name', 'File Type', 'File Url', 'Upload Date', 'Change Date', 'Contents', 'aiscore関数', 'メールアドレス', '送信チェック']);
  } else {
    // 既存のヘッダーを更新
    const headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
    if (headers[7] !== 'aiscore関数' || headers[8] !== 'メールアドレス') {
      sheet.getRange(1, 8, 1, 2).setValues([['aiscore関数', 'メールアドレス']]);
    }
  }

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const fullFileName = file.getName();
    
    // ファイル名を処理
    const fileName = fullFileName.includes(" - ") ? fullFileName.split(" - ")[0] : fullFileName;
    
    const fileMimeType = file.getMimeType();
    const fileUrl = file.getUrl();
    const uploadDate = file.getDateCreated();
    const ownerEmail = file.getOwner().getEmail();
    const newRowIndex = sheet.getLastRow() + 1;

    const existingRowIndex = findRowByFileId(sheet, fileId);
    if (existingRowIndex > 1) {  // ヘッダー行を除外
      const changeDate = new Date();
      sheet.getRange(existingRowIndex, 2).setValue(fileName);
      sheet.getRange(existingRowIndex, 6).setValue(changeDate);
      // sheet.getRange(existingRowIndex, 9).setValue(ownerEmail);
      getFileContent(file, sheet, existingRowIndex);
    } else {
      // sheet.appendRow([fileId, fileName, fileMimeType, fileUrl, uploadDate, '', '', '', ownerEmail]);
      sheet.appendRow([fileId, fileName, fileMimeType, fileUrl, uploadDate, '', '', '',  '']);
      getFileContent(file, sheet, newRowIndex);
    }
  }

  // ドライブから削除されたファイルがあれば、スプレッドシートから対応する行を削除
  const allFiles = folder.getFiles();
  const fileIdsInDrive = [];
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const fileId = file.getId();
    fileIdsInDrive.push(fileId);
  }

  for (let i = 2; i <= sheet.getLastRow(); i++) {  // ヘッダー行をスキップ
    const fileId = sheet.getRange(i, 1).getValue();
    if (!fileIdsInDrive.includes(fileId)) {
      sheet.deleteRow(i);
      i--;
    }
  }

  // フォルダを一般公開にする
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ログにも完了メッセージを出力
  Logger.log("データの読込みがすべて完了しました");

  // 完了メッセージをメッセージボックスで表示
  Browser.msgBox("処理完了", "データの読込みがすべて完了しました", Browser.Buttons.OK);
}

// ファイルIDに基づいてスプレッドシート内の行を見つける関数
function findRowByFileId(sheet, fileId) {
  if (!sheet) {
    Logger.log(`Error finding row by file ID: sheet is undefined.`);
    return -1;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  for (let i = 1; i < data.length; i++) {  // ヘッダー行をスキップ
    if (data[i][0] === fileId) {
      return i + 1;  // スプレッドシートの行番号は1から始まるため
    }
  }
  return -1;
}

// スプレッドシートに保存されているすべてのファイルIDを取得する関数
function getAllFileIds(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const fileIds = [];
  for (let i = 1; i < data.length; i++) {  // ヘッダー行をスキップ
    fileIds.push(data[i][0]);
  }
  return fileIds;
}

// ファイルの内容を取得し、スプレッドシートに書き込む関数
// Google Colabファイルを抽出してスプレッドシートに表示する関数
function getFileContent(file, sheet, rowIndex) {
  if (!file || !sheet) {
    Logger.log(`Error getting file content: file or sheet is undefined.`);
    return "";
  }

  const fileMimeType = file.getMimeType();

  if (fileMimeType === 'application/vnd.google.colaboratory') {
    const lastModifierEmail = getFileLastModifierEmail(file.getId());
    getColabNotebookContent(file, sheet, rowIndex, lastModifierEmail); // Colabノートブックの内容を取得して書き込み
  } else {
    // 他のファイルタイプの場合の処理
    const lastModifierEmail = getFileLastModifierEmail(file.getId());
    if (fileMimeType === MimeType.GOOGLE_DOCS) {
      return getGoogleDocsContent(file, sheet, rowIndex, lastModifierEmail);
    } else if (fileMimeType === MimeType.PLAIN_TEXT) {
      return getTextContent(file, sheet, rowIndex, lastModifierEmail);
    } else if (fileMimeType === 'application/json') {
      return getJsonContent(file, sheet, rowIndex, lastModifierEmail);
    } else if (fileMimeType === 'application/pdf') {
      return getPdfContent(file, sheet, rowIndex, lastModifierEmail);
    } else if (fileMimeType === MimeType.MICROSOFT_WORD || 
               fileMimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
               fileMimeType === 'application/vnd.ms-word.document.macroenabled.12' ||
               fileMimeType === 'application/msword') {
      return getWordContent(file, sheet, rowIndex, lastModifierEmail);
    } else {
      Logger.log(`Unsupported file type: ${fileMimeType}`);
      return "";
    }
  }
}

function getFileLastModifierEmail(fileId) {
  try {
    var file = Drive.Files.get(fileId, {fields: 'lastModifyingUser'});
    if (file.lastModifyingUser && file.lastModifyingUser.emailAddress) {
      return file.lastModifyingUser.emailAddress;
    }
  } catch (e) {
    Logger.log(`Error getting file last modifier email: ${e.message}`);
  }
  return "Unknown";
}

function getGoogleDocsContent(file, sheet, rowIndex, lastModifierEmail) {
  try {
    const doc = DocumentApp.openById(file.getId());
    const text = doc.getBody().getText();
    doc.saveAndClose();
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), text, lastModifierEmail);
    return text;
  } catch (e) {
    Logger.log(`Error getting Google Docs content for file ${file.getName()}: ${e.message}`);
    return "";
  }
}

function getTextContent(file, sheet, rowIndex, lastModifierEmail) {
  if (!file) {
    Logger.log(`Error getting text content: file is undefined.`);
    return "";
  }
  try {
    let text = file.getBlob().getDataAsString('UTF-8');
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), text, lastModifierEmail);
    return text;
  } catch (e) {
    if (file && file.getName) {
      Logger.log(`Error getting text content for file ${file.getName()}: ${e.message}`);
    } else {
      Logger.log(`Error getting text content: ${e.message}`);
    }
    return "";
  }
}

function getJsonContent(file, sheet, rowIndex, lastModifierEmail) {
  try {
    const text = file.getBlob().getDataAsString();
    const json = JSON.parse(text);
    const formattedText = JSON.stringify(json, null, 2);
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), formattedText, lastModifierEmail);
    return formattedText;
  } catch (e) {
    if (file && file.getName) {
      Logger.log(`Error getting JSON content for file ${file.getName()}: ${e.message}`);
    } else {
      Logger.log(`Error getting JSON content: ${e.message}`);
    }
    return "";
  }
}

function getPdfContent(file, sheet, rowIndex, lastModifierEmail) {
  try {
    const blob = file.getBlob();
    const tempFolder = DriveApp.createFolder('TempFolder_' + new Date().getTime());
    const tempFile = tempFolder.createFile(blob);
    
    // OCRを使用してGoogle Docsに変換
    const ocrFile = Drive.Files.copy(
      { title: file.getName(), mimeType: MimeType.GOOGLE_DOCS },
      tempFile.getId(),
      { ocr: true, ocrLanguage: 'ja' }  // 日本語OCRを指定
    );
    
    const doc = DocumentApp.openById(ocrFile.id);
    const text = doc.getBody().getText();
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), text, lastModifierEmail);
    
    // 一時ファイルとフォルダを削除
    DriveApp.getFileById(ocrFile.id).setTrashed(true);
    tempFile.setTrashed(true);
    tempFolder.setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log(`Error getting PDF content for file ${file.getName()}: ${e.message}`);
    return "";
  }
}

function getWordContent(file, sheet, rowIndex, lastModifierEmail) {
  try {
    const blob = file.getBlob();
    const tempFolder = DriveApp.createFolder('TempFolder_' + new Date().getTime());
    const tempFile = tempFolder.createFile(blob);
    
    // Google Docsに変換
    const mimeType = MimeType.GOOGLE_DOCS;
    const convertedFile = Drive.Files.copy(
      { title: file.getName(), mimeType: mimeType },
      tempFile.getId()
    );
    
    const doc = DocumentApp.openById(convertedFile.id);
    const text = doc.getBody().getText();
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), text, lastModifierEmail);
    
    // 一時ファイルとフォルダを削除
    DriveApp.getFileById(convertedFile.id).setTrashed(true);
    tempFile.setTrashed(true);
    tempFolder.setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log(`Error getting Word content for file ${file.getName()}: ${e.message}`);
    return "";
  }
}

// Colabノートブックファイルの内容を取得してスプレッドシートに書き込む関数
function getColabNotebookContent(file, sheet, rowIndex, lastModifierEmail) {
  try {
    // ファイルの内容を取得し、JSON形式としてパース
    const jsonContent = JSON.parse(file.getBlob().getDataAsString());
    const cells = jsonContent.cells;

    let textContent = ''; // 抽出されたコードやMarkdownの内容を保持する変数

    // 各セルの内容を抽出
    for (let i = 0; i < cells.length; i++) {
      if (cells[i].cell_type === 'code') {
        if (cells[i].source.join('')===''){
          textContent += 'コードなし\n';
        } else {
          textContent += 'コード\n' + cells[i].source.join('') + '\n\n';
        }
      } else if (cells[i].cell_type === 'markdown') {
        textContent += cells[i].source.join('') + '\n';
      }
    }

    // スプレッドシートに書き込む
    writeTextToSheet(sheet, rowIndex, file.getId(), file.getName(), file.getMimeType(), file.getUrl(), file.getDateCreated(), new Date(), textContent, lastModifierEmail);
    
    return textContent; // すべてのセルの内容を返す
  } catch (e) {
    Logger.log('Error extracting Colab file content: ' + e.message);
    return 'Error extracting content';
  }
}

// スプレッドシートにデータを書き込む関数
function writeTextToSheet(sheet, rowIndex, fileId, fullFileName, fileMimeType, fileUrl, uploadDate, changeDate, text, lastModifierEmail) {
  const fileName = fullFileName.includes(" - ") ? fullFileName.split(" - ")[0] : fullFileName;

  if (text) {
    const encodedText = Utilities.newBlob(text, 'text/plain', 'temp.txt').getDataAsString('UTF-8');
    const maxCellLength = 50000;
    if (encodedText.length > maxCellLength) {
      const rowsCount = Math.ceil(encodedText.length / maxCellLength);
      for (let i = 0; i < rowsCount; i++) {
        const start = i * maxCellLength;
        const end = Math.min((i + 1) * maxCellLength, encodedText.length);
        if (i === 0) {
          sheet.getRange(rowIndex, 1).setValue(fileId);
          sheet.getRange(rowIndex, 2).setValue(fileName);
          sheet.getRange(rowIndex, 3).setValue(fileMimeType);
          sheet.getRange(rowIndex, 4).setValue(fileUrl);
          sheet.getRange(rowIndex, 5).setValue(uploadDate);
          sheet.getRange(rowIndex, 6).setValue(changeDate);
          sheet.getRange(rowIndex, 9).setValue(lastModifierEmail);
        } else {
          sheet.insertRowAfter(rowIndex + i - 1);
        }
        sheet.getRange(rowIndex + i, 7).setValue(encodedText.substring(start, end));
      }
    } else {
      sheet.getRange(rowIndex, 1).setValue(fileId);
      sheet.getRange(rowIndex, 2).setValue(fileName);
      sheet.getRange(rowIndex, 3).setValue(fileMimeType);
      sheet.getRange(rowIndex, 4).setValue(fileUrl);
      sheet.getRange(rowIndex, 5).setValue(uploadDate);
      sheet.getRange(rowIndex, 6).setValue(changeDate);
      sheet.getRange(rowIndex, 7).setValue(encodedText);
      sheet.getRange(rowIndex, 9).setValue(lastModifierEmail);
    }
    // aiscore関数列は空のままにしておく
    sheet.getRange(rowIndex, 8).setValue('');
  }
}
