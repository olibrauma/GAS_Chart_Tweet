var service        = twitter.getService();
var endPointStatus = 'https://api.twitter.com/1.1/statuses/update.json';
var endPointMedia  = 'https://upload.twitter.com/1.1/media/upload.json';
 
 
function mediaTweet(){ 
 
  //ツイート本文と画像のURL  
  var twText = "This is a sample.";
  var imgUrl = "https://sample.com/sample.jpg";
     
  //画像の取得
  var imgBlob = UrlFetchApp.fetch(imgUrl).getBlob();
  var img_64  = Utilities.base64Encode(imgBlob.getBytes());
   
  var img_option = { 
    'method' : "POST", 
    'payload': {
      'media_data': img_64
    } 
  };
   
   
  var image_upload = JSON.parse(service.fetch(endPointMedia, img_option)); 
   
  var sendoption = { 
    'status'   : twText, 
    'media_ids': image_upload['media_id_string']
  };
 
  service.fetch(endPointStatus, {method: 'post', payload: sendoption});
}
