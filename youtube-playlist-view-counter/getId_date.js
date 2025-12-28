// プレイリストから動画のIDと公開日を取得
// https://developers.google.com/youtube/v3/docs/playlistItems/list
function getId_date(params, videoCount){
   
  var url = "https://www.googleapis.com/youtube/v3/playlistItems";
   
  url += "?part="       + "contentDetails";
  url += "&playlistId=" + params.playlistId;
  url += "&maxResults=" + params.maxResults;
  url += "&key="        + params.key;
   
  // APIレスポンスを入れる変数
  var data = [];
   
  // 2回目以降のAPIリクエスト時に，次のページを要求するためのトークン
  var pageToken = "";
   
  // 動画の数だけデータをAPIリクエストを出す
  do{
    // ページトークンの有無によってURLをちょっと変える
    var req = url;
    if(pageToken != "") req += "&pageToken=" + pageToken;
     
    // APIリクエストを変数に入れる
    var res = JSON.parse(UrlFetchApp.fetch(req));
     
    // リクエストデータから，IDと公開日を取り出す
    for (var i in res.items){
      data.push(res.items[i].contentDetails);
    }
     
    // 次のページのトークンがあれば更新
    if(res.nextPageToken != "") pageToken = res.nextPageToken;
     
  } while (data.length < videoCount);
   
  return data;
}
