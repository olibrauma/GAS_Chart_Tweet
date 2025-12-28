// 動画のIDを maxResults 個ずつに分割して，それぞれをコンマ連結のテキストにする
function makeIdList(params, id_date){
     
  var length = id_date.length;              // 動画の数
  var loops = Math.ceil(length/params.maxResults); // 外側ループの回数．動画が130個なら3回ループが要る
  var idList = [];                          // 出力用の変数
   
  // 外側ループ
  for (var i=0; i < loops; i++){
     
    // IDのコンマ連結を入れる変数
    var idText = "";
     
    // maxResults 個ずつに分けて動画IDのコンマ連結を作る
    for (var count=0; count < params.maxResults; count++){
       
      // コンマ連結を作る
      try { idText += id_date[params.maxResults*i + count].videoId + ","; }
       
      // 最後のループ (i = loops-1) では maxResults 回も繰り返す必要がないので，その場合はループを抜ける
      // 例えば maxResults=50 なら，
      // 0<id<49 (i=0), 50<id<99 (i=1), 100<id<149 (i=2), ... という感じ
      catch(e) { break; }
    }
     
    // 末尾に不要なコンマを付けているので，削除してリストに追加
    idList.push(idText.slice(0, -1));
  }
   
  return idList;
}
