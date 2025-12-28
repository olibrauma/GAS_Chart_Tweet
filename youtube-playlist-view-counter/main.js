// 動画データ = id, date, view のリストを作る
function getVideoData(){
   
  // APIリクエストごとに変わらない値を入れておく
  var params = {
    maxResults : 50,
    playlistId : "YOUR-PLAYLIST-ID-HERE",
    key        : "YOUR-API-KEY-HERE"
  }
   
  // 動画の総数を取得
  var videoCount = getVideoCount(params);
   
  // 動画のIDと公開日のデータを取得
  var id_date = getId_date(params, videoCount);
   
  // 取得した再生回数と id_date を結合して出力用の変数に入れる
  var videoData = getViewCount(params, id_date);
   
  return videoData;
}
