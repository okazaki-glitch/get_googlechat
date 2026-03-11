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
    ui.alert(buildChatApiErrorMessage_(error));
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
      orderBy: 'createTime ASC',
      pageSize: 1000
    };

    // スレッド指定がある場合のみフィルターする。指定がない場合はスペース全体を取得する。
    if (threadName) {
      options.filter = `thread.name = ${threadName}`;
    }

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
    ['スレッド名', threadInfo.threadName || '（指定なし: スペース全体）']
  ]);

  const headerRow = 6;
  sheet.getRange(headerRow, 1, 1, 6).setValues([
    ['投稿日時', '投稿者ID', '本文', 'メッセージID', '画像/添付URL', 'オブジェクトURL']
  ]);

  if (messages.length === 0) {
    sheet.getRange(headerRow + 1, 1).setValue('対象スレッドにメッセージがありません。');
    return;
  }

  const rows = messages.map((message) => {
    const createTime = message.createTime
      ? Utilities.formatDate(new Date(message.createTime), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : '';
    const sender = extractSenderId_(message.sender);
    const text = extractMessageText_(message);
    const messageId = message.name || '';
    const attachmentUrls = collectAttachmentAndImageUrls_(message);
    const objectUrls = collectObjectUrls_(message);

    return [createTime, sender, text, messageId, attachmentUrls, objectUrls];
  });

  sheet.getRange(headerRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, 6);
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

function collectAttachmentAndImageUrls_(message) {
  const urls = [];
  const attachments = []
    .concat(Array.isArray(message.attachment) ? message.attachment : [])
    .concat(Array.isArray(message.attachments) ? message.attachments : []);

  attachments.forEach((attachment) => {
    if (attachment.downloadUri) {
      urls.push(attachment.downloadUri);
    }
    if (attachment.thumbnailUri) {
      urls.push(attachment.thumbnailUri);
    }
    const driveFileId = attachment.driveDataRef && attachment.driveDataRef.driveFileId;
    if (driveFileId) {
      urls.push(buildDriveFileUrlFromId_(driveFileId));
    }
  });

  const gifs = Array.isArray(message.attachedGifs) ? message.attachedGifs : [];
  gifs.forEach((gif) => {
    if (gif.uri) {
      urls.push(gif.uri);
    }
  });

  return normalizeAndJoinUrls_(urls);
}

function collectObjectUrls_(message) {
  const urls = [];

  if (message.matchedUrl && message.matchedUrl.url) {
    urls.push(message.matchedUrl.url);
  }

  const annotations = Array.isArray(message.annotations) ? message.annotations : [];
  annotations.forEach((annotation) => {
    const richLink = annotation.richLinkMetadata;
    if (richLink && richLink.uri) {
      urls.push(richLink.uri);
    }
    const driveFileId =
      richLink &&
      richLink.driveLinkData &&
      richLink.driveLinkData.driveDataRef &&
      richLink.driveLinkData.driveDataRef.driveFileId;
    if (driveFileId) {
      urls.push(buildDriveFileUrlFromId_(driveFileId));
    }
  });

  // Chatアプリのカード/ウィジェット内URLも抽出する。
  const cardLikeParts = [message.cards, message.cardsV2, message.accessoryWidgets];
  cardLikeParts.forEach((part) => {
    urls.push(...extractUrlsFromNestedObject_(part));
  });

  return normalizeAndJoinUrls_(urls);
}

function extractUrlsFromNestedObject_(value) {
  const urls = [];

  function traverse_(target) {
    if (!target) {
      return;
    }
    if (Array.isArray(target)) {
      target.forEach((item) => traverse_(item));
      return;
    }
    if (typeof target !== 'object') {
      return;
    }

    Object.keys(target).forEach((key) => {
      const child = target[key];
      if (typeof child === 'string' && isLikelyUrlFieldKey_(key) && /^https?:\/\//i.test(child)) {
        urls.push(child);
        return;
      }
      traverse_(child);
    });
  }

  traverse_(value);
  return urls;
}

function isLikelyUrlFieldKey_(key) {
  return /(^url$|^uri$|imageUrl$|thumbnailUri$|downloadUri$)/i.test(key);
}

function buildDriveFileUrlFromId_(driveFileId) {
  return `https://drive.google.com/file/d/${driveFileId}/view`;
}

function normalizeAndJoinUrls_(urls) {
  const uniqueUrls = [];
  const seen = {};

  urls.forEach((url) => {
    if (!url) {
      return;
    }
    const normalized = String(url).trim();
    if (!normalized || seen[normalized]) {
      return;
    }
    seen[normalized] = true;
    uniqueUrls.push(normalized);
  });

  return uniqueUrls.join('\n');
}

function extractSenderId_(sender) {
  if (!sender) {
    return '';
  }
  const senderName = sender.name || '';
  if (!senderName) {
    return '';
  }

  // 例: users/12345678901234567890 -> 12345678901234567890
  const userMatch = senderName.match(/^users\/([^\/]+)$/);
  if (userMatch) {
    return userMatch[1];
  }

  // users以外(apps/..., anonymousUsers/...)は識別のため元値をそのまま返す。
  return senderName;
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
    const roomMatch = candidate.match(/\/room\/([^\/?#&]+)(?:\/([^\/?#&]+))?/i);
    if (roomMatch) {
      const spaceId = roomMatch[1];
      const threadId = roomMatch[2];
      return threadId
        ? {
            spaceName: `spaces/${spaceId}`,
            threadName: `spaces/${spaceId}/threads/${threadId}`
          }
        : {
            spaceName: `spaces/${spaceId}`,
            threadName: ''
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

  // mail.google.com/.../#chat/space/{space} 形式を解析する（スレッド指定なし）。
  for (const candidate of candidates) {
    const spaceOnlyMatch = candidate.match(/\/space\/([^\/?#&]+)/i);
    if (spaceOnlyMatch) {
      const spaceId = spaceOnlyMatch[1];
      return {
        spaceName: `spaces/${spaceId}`,
        threadName: ''
      };
    }
  }

  throw new Error(
    'スレッドURLからスペースIDとスレッドIDを抽出できませんでした。' +
      'URLに「spaces/.../threads/...」「.../space/{space}/thread/{thread}」または「.../space/{space}」が含まれているか確認してください。'
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

function buildChatApiErrorMessage_(error) {
  const rawMessage = error && error.message ? String(error.message) : String(error);

  if (rawMessage.includes('Google Chat app not found')) {
    return (
      '取得に失敗しました: Google Chat APIのアプリ設定が未完了です。\n\n' +
      '対応手順:\n' +
      '1) Apps Scriptの「プロジェクトの設定」で、標準のGoogle Cloudプロジェクトを紐付ける\n' +
      '2) そのGoogle Cloudプロジェクトで「Google Chat API」を有効化する\n' +
      '3) Google Chat API > Configurationでアプリ情報（名前/アイコンURL/説明）を入力して保存する\n' +
      '4) スクリプトを再実行して権限を再承認する\n\n' +
      '補足: Chat APIはGoogle Workspaceアカウントが必要です。'
    );
  }

  return `取得に失敗しました: ${rawMessage}`;
}
