function getChartBase64(){
   
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("graph");
  var chart = sheet.getCharts()[0];
   
  chart = chart.modify().build();
  sheet.updateChart(chart);  
   
  return imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes());
 
}
