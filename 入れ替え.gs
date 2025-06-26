function myFunction() {

    const sheet = SpreadsheetApp.getActiveSheet();
     var lastRow = sheet.getLastRow();
    for (var i = 2; i < lastRow+1; i++){
      var usageDate=sheet.getRange(i,1).getValue();
      console.log(usageDate);
       var usageDate = Utilities.formatDate(usageDate, 'JST', 'yyyy-MM-dd');
             console.log(usageDate);
      var comment=sheet.getRange(i,2).getValue();
      var amount=sheet.getRange(i,3).getValue();
      if (comment=="物販")
      {
         shop="物販";
         comment="システムから入力";

        var category_id =102;
        var genre_id = 10201;

      }
      else if(comment=="バスチャージ 車載端末"){
        var category_id =199;
        var genre_id = 19908;



      }

    }
}
