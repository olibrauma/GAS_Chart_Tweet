// data は数値の配列
function calcStatsSummary(data){
   
  // データ数 (サンプルサイズ) を取得
  var size = data.length;
   
  // 平均を計算
  var sum1 = 0;
  for (var i in data){
    sum1 += data[i];
  }
  var mean = sum1/size;
   
  // (不偏) 標準偏差を計算
  var sum2 = 0;
  for (var i in data){
    sum2 += Math.pow(data[i] - mean, 2);
  }
  var SD = Math.sqrt(sum2/(size - 1));
   
  // 歪度と尖度を計算
  var sum3 = 0;
  var sum4 = 0;
  for (var i in data){
    sum3 += Math.pow((data[i] - mean)/SD, 3);
    sum4 += Math.pow((data[i] - mean)/SD, 4);
  }
  var skew = sum3*size/(size-1)/(size-2);
  var exKurt = size*(size+1)/(size-1)/(size-2)/(size-3)*sum4 - 3*(size-1)*(size-1)/(size-2)/(size-3);
   
  // 各種代表値が返り値
  return {size: size, mean: mean, SD: SD, skew: skew, exKurt: exKurt};
}
