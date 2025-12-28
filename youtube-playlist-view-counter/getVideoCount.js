// プレイリスト内の動画の数を取得
// https://developers.google.com/youtube/v3/docs/playlists
function getVideoCount(params){
 
  var url = "https://www.googleapis.com/youtube/v3/playlists";
   
  url += "?part=" + "contentDetails";
  url += "&id="   + params.playlistId;
  url += "&key="  + params.key;
   
  var res = JSON.parse(UrlFetchApp.fetch(url));
   
  return res.items[0].contentDetails.itemCount;
}
