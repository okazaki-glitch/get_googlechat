const RUN_COUNT_PROPERTY_KEY = 'chatThreadExportRunCount';

function copyGoogleChatThreadToSheet() {
  const ui = SpreadsheetApp.getUi();
  const promptResult = ui.prompt(
    'Google Chatスレッドのエクスポート',
    'コピーしたいGoogle ChatスレッドのURLを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const threadUrl = promptResult.getResponseText().trim();
  if (!threadUrl) {
    ui.alert('URLが空です。Google ChatスレッドのURLを入力してください。');
    return;
  }

  let threadInfo;
  try {
    threadInfo = parseThreadFromUrl_(threadUrl);
  } catch (error) {
    ui.alert(error.message);
    return;
  }

  try {
    const messages = fetchThreadMessages_(threadInfo.spaceName, threadInfo.threadName);
    const targetSheet = getTargetSheet_();
    writeMessagesToSheet_(targetSheet, threadUrl, threadInfo, messages);

    ui.alert(`完了: ${messages.length}件のメッセージを「${targetSheet.getName()}」へ出力しました。`);
  } catch (error) {
    ui.alert(`取得に失敗しました: ${error.message}`);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Chatエクスポート')
    .addItem('スレッドをシートへコピー', 'copyGoogleChatThreadToSheet')
    .addToUi();
}

function fetchThreadMessages_(spaceName, threadName) {
  const messages = [];
  let pageToken = null;

  do {
    const options = {
      filter: `thread.name = ${threadName}`,
      orderBy: 'createTime ASC',
      pageSize: 1000
    };

    if (pageToken) {
      options.pageToken = pageToken;
    }

    const response = Chat.Spaces.Messages.list(spaceName, options);

    if (response.messages && response.messages.length > 0) {
      messages.push(...response.messages);
    }

    pageToken = response.nextPageToken || null;
  } while (pageToken);

  return messages;
}

function writeMessagesToSheet_(sheet, threadUrl, threadInfo, messages) {
  sheet.clear();

  const now = new Date();
  sheet.getRange(1, 1, 4, 2).setValues([
    ['取得日時', Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')],
    ['スレッドURL', threadUrl],
    ['スペース名', threadInfo.spaceName],
    ['スレッド名', threadInfo.threadName]
  ]);

  const headerRow = 6;
  sheet.getRange(headerRow, 1, 1, 4).setValues([
    ['投稿日時', '投稿者', '本文', 'メッセージID']
  ]);

  if (messages.length === 0) {
    sheet.getRange(headerRow + 1, 1).setValue('対象スレッドにメッセージがありません。');
    return;
  }

  const rows = messages.map((message) => {
    const createTime = message.createTime
      ? Utilities.formatDate(new Date(message.createTime), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : '';
    const sender = message.sender && message.sender.displayName ? message.sender.displayName : '';
    const text = extractMessageText_(message);
    const messageId = message.name || '';

    return [createTime, sender, text, messageId];
  });

  sheet.getRange(headerRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, 4);
}

function extractMessageText_(message) {
  if (message.text) {
    return message.text;
  }
  if (message.formattedText) {
    return message.formattedText;
  }
  return '';
}

function parseThreadFromUrl_(threadUrl) {
  const decodedUrl = decodePossiblyEncodedUrl_(threadUrl);
  const candidates = [threadUrl, decodedUrl];

  // URL中にAPI形式のリソース名が含まれるケースを優先して抽出する。
  for (const candidate of candidates) {
    const threadResourceMatch = candidate.match(/(spaces\/[^\/?#&]+\/threads\/[^\/?#&]+)/i);
    if (threadResourceMatch) {
      const threadName = threadResourceMatch[1];
      const spaceName = threadName.match(/spaces\/[^\/?#&]+/i)[0];
      return {spaceName: normalizeResourceName_(spaceName), threadName: normalizeResourceName_(threadName)};
    }
  }

  // chat.google.com/room/{space}/{thread} 形式を解析する。
  for (const candidate of candidates) {
    const roomMatch = candidate.match(/\/room\/([^\/?#&]+)\/([^\/?#&]+)/i);
    if (roomMatch) {
      const spaceId = roomMatch[1];
      const threadId = roomMatch[2];
      return {
        spaceName: `spaces/${spaceId}`,
        threadName: `spaces/${spaceId}/threads/${threadId}`
      };
    }
  }

  // mail.google.com/chat/u/0/#chat/space/{space}/thread/{thread} 形式を解析する。
  for (const candidate of candidates) {
    const hashThreadMatch = candidate.match(/\/space\/([^\/?#&]+)\/thread\/([^\/?#&]+)/i);
    if (hashThreadMatch) {
      const spaceId = hashThreadMatch[1];
      const threadId = hashThreadMatch[2];
      return {
        spaceName: `spaces/${spaceId}`,
        threadName: `spaces/${spaceId}/threads/${threadId}`
      };
    }
  }

  throw new Error(
    'スレッドURLからスペースIDとスレッドIDを抽出できませんでした。' +
      'URLに「spaces/.../threads/...」または「.../space/{space}/thread/{thread}」が含まれているか確認してください。'
  );
}

function decodePossiblyEncodedUrl_(url) {
  let decoded = url;

  // 2回までデコードして、エンコード済みのリソース名を抽出しやすくする。
  for (let i = 0; i < 2; i += 1) {
    try {
      const next = decodeURIComponent(decoded);
      if (next === decoded) {
        break;
      }
      decoded = next;
    } catch (error) {
      break;
    }
  }

  return decoded;
}

function normalizeResourceName_(resourceName) {
  return resourceName.replace(/^\/+/, '');
}

function getTargetSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const runCount = Number(scriptProperties.getProperty(RUN_COUNT_PROPERTY_KEY) || '0');

  let sheet;
  if (runCount === 0) {
    sheet = spreadsheet.getActiveSheet();
  } else {
    sheet = spreadsheet.insertSheet(buildNewSheetName_(spreadsheet));
  }

  scriptProperties.setProperty(RUN_COUNT_PROPERTY_KEY, String(runCount + 1));
  return sheet;
}

function buildNewSheetName_(spreadsheet) {
  const baseName = `Chat_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
  let sheetName = baseName;
  let suffix = 1;

  while (spreadsheet.getSheetByName(sheetName)) {
    sheetName = `${baseName}_${suffix}`;
    suffix += 1;
  }

  return sheetName;
}
