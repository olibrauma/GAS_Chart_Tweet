function fetchYouTubeDataForIFTT() {
  // --- 設定 ---
  const SHEET_NAME = 'IFTTT';
  const COL_URL = 1;      // B列 (配列のインデックスは1)
  const COL_TITLE = 0;    // A列
  const COL_DATE = 2;     // C列
  const COL_DURATION = 6; // G列
  // -----------

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // データがない場合は終了

  // 2行目から最終行までの範囲を取得
  const range = sheet.getRange(2, 1, lastRow - 1, 7);
  const values = range.getValues();

  // 更新が必要な行のリストを作成
  let updates = []; 
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const url = row[COL_URL];
    
    // URLがない、または 全て埋まっている場合はスキップ
    if (!url || (row[COL_TITLE] && row[COL_DATE] && row[COL_DURATION])) {
      continue;
    }

    const videoId = extractVideoId(url);
    if (videoId) {
      updates.push({ rowIndex: i, videoId: videoId });
    }
  }

  if (updates.length === 0) {
    console.log('更新対象の行はありませんでした。');
    return;
  }

  // YouTube Data API (50件ずつ処理)
  const CHUNK_SIZE = 50;
  for (let i = 0; i < updates.length; i += CHUNK_SIZE) {
    const chunk = updates.slice(i, i + CHUNK_SIZE);
    const ids = chunk.map(u => u.videoId).join(',');

    try {
      const response = YouTube.Videos.list('snippet,contentDetails', { id: ids });
      
      if (!response.items) continue;

      const videoMap = {};
      response.items.forEach(item => {
        videoMap[item.id] = {
          title: item.snippet.title,
          publishedAt: item.snippet.publishedAt.substr(0, 10), // YYYY-MM-DD
          duration: parseDuration(item.contentDetails.duration)
        };
      });

      chunk.forEach(u => {
        const data = videoMap[u.videoId];
        
        if (data) {
          const targetRow = u.rowIndex + 2;
          
          // タイトル
          if (data.title) {
            sheet.getRange(targetRow, COL_TITLE + 1).setValue(data.title);
          }
          
          // 公開日
          if (data.publishedAt) {
            sheet.getRange(targetRow, COL_DATE + 1).setValue(data.publishedAt);
          }
          
          // 動画時間 (parseDurationの結果がnullでない場合のみ記入)
          if (data.duration) { 
            sheet.getRange(targetRow, COL_DURATION + 1).setValue(data.duration);
          }
        }
      });
    
    } catch (e) {
      console.error('APIリクエストエラー: ' + e.toString());
    }
  }
  
  console.log('処理が完了しました。');
}

/**
 * YouTubeのURLから動画IDを抽出する
 */
function extractVideoId(url) {
  if (!url) return null;
  const strUrl = url.toString();
  const match = strUrl.match(/(?:v=|\/)([0-9A-Za-z_-]{11})/);
  return match ? match[1] : null;
}

/**
 * ISO 8601形式 (PT1H2M10S) を 00:mm:ss 形式に変換する
 * ⚠️ 変換できない場合は null を返します
 */
function parseDuration(duration) {
  // duration自体がnull/undefinedの場合に null を返す
  if (!duration) {
    return null;
  }

  const match = duration.match(/PT(\d+H)?(\d+M)?(\d+S)?/);
  
  // 正規表現がマッチしなかった場合（予期せぬ形式やライブ動画など）に null を返す
  if (!match) {
    return null; 
  }

  const hours = (match[1] || '').replace('H', '');
  const minutes = (match[2] || '').replace('M', '');
  const seconds = (match[3] || '').replace('S', '');

  const h = parseInt(hours) || 0;
  const m = parseInt(minutes) || 0;
  const s = parseInt(seconds) || 0;

  // ゼロ埋め関数
  const pad = (num) => num.toString().padStart(2, '0');

  // 常に HH:mm:ss の形式で返す
  return `${pad(h)}:${pad(m)}:${pad(s)}`;
}

/**
 * スプレッドシートが開かれたときにカスタムメニューを追加する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ YouTube情報取得')
      .addItem('動画情報を更新 (B列から取得)', 'fetchYouTubeDataForIFTT')
      .addToUi();
}