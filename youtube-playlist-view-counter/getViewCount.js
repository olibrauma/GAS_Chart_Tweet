// 取得した再生回数と id_date を結合して出力用の変数に入れる
function getViewCount(params, id_date){
   
  // id のカンマ区切りリストを作る
  var idList = makeIdList(params, id_date);
   
  // 出力用の変数
  var videoData = [];
   
  // maxResults 個ずつAPIリクエストでデータを取得
  for (var i in idList){
     
    // APIのエンドポイントとパラメータ
    var url = "https://www.googleapis.com/youtube/v3/videos";
   
    url += "?part="   + "statistics";
    url += "&id="     + idList[i];
    url += "&fields=" + "items%2Fstatistics";
    url += "&key="    + params.key;
   
    // idList に ID が含まれる動画のデータを取得
    var res = JSON.parse(UrlFetchApp.fetch(url));
     
    // 取得した再生回数と id_date を結合して出力用の変数に入れる
    for (var j in res.items) {
       
      var stats = {
        id   : id_date[Number(i)*params.maxResults + Number(j)].videoId, 
        date : id_date[Number(i)*params.maxResults + Number(j)].videoPublishedAt.slice(0, 10),
        view : Number(res.items[j].statistics.viewCount)
      }
     
      videoData.push(stats);
    }
  }
   
  return videoData;
}
