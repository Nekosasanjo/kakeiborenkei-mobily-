function seikikamobi() {
    var spreadsheet = SpreadsheetApp.getActiveSheet();
    var lastRow=spreadsheet.getLastRow();
    var range = spreadsheet.getRange('E1');
    var values = range.setValue('利用会社');
    var kari=spreadsheet.getRange("F1");
    var kengen=kari.setValue("編集権限");


    for (var i = 2; i < lastRow+1; i++){
      shubetu=spreadsheet.getRange(i,4).getValue();
      kari=spreadsheet.getRange(i,6).getValue();
      useday=spreadsheet.getRange(i,1).getValue();
      money=spreadsheet.getRange(i,3).getValue();
      lenmoney=money.length;
      basho=spreadsheet.getRange(i,2).getValue();
      kaigyo=0;
      for (var j=0;j<basho.length;j++){
        if (basho[j]=="\n"){
          kaigyo=j;
        }
      }

      bashomae=basho.slice(0,kaigyo);
      bashoato=basho.slice(kaigyo+1,basho.length);
      if (basho!="" && basho[0]!="入"){
        riyou="入  "+bashomae+"\n"+"出  "+bashoato;
        var are=spreadsheet.getRange(i,2).setValue(riyou);
      }


      if (lenmoney>=10)
      {
        if (lenmoney==10)
        {
          money=money.slice(6,10)
          var are=spreadsheet.getRange(i,3).setValue(money);
        }
        else if (lenmoney==12)
        {
          money=money.slice(8,13)
          var are=spreadsheet.getRange(i,3).setValue(money);
        }

      }



      var nowuseday=spreadsheet.getRange(i,1).getDisplayValue();
      if (nowuseday.indexOf(".")==-1)
      {
        var haifun="kaenai";
      }
      else{
        var haifun="kaeru";
      }

      if (haifun=="kaeru")
      {
        for(var j=1;j<=3;j++){
          useday=useday.replace(".","-");
        }
      }


      useday=spreadsheet.getRange(i,1).setValue(useday);
      if (shubetu=="乗降"&& kari=="")
      {
        spreadsheet.getRange(i,1,2,1)
        .breakApart();
        spreadsheet.getRange(i+1,2,1,3)
        .breakApart();
        var kaisha=spreadsheet.getRange(i+1,2).getValue();        
        var range = spreadsheet.getRange(i,5);
        var values = range.setValue(kaisha);
        spreadsheet.deleteRow(i+1);
        kari=spreadsheet.getRange(i,6);
        kengen=kari.setValue("look")
      }
      if (shubetu=="乗降乗越精算"&& kari=="")
      {
        spreadsheet.getRange(i,1,2,1)
        .breakApart();
        spreadsheet.getRange(i+1,2,1,3)
        .breakApart();
        var kaisha=spreadsheet.getRange(i+1,2).getValue();        
        var range = spreadsheet.getRange(i,5);
        var values = range.setValue(kaisha);
        spreadsheet.deleteRow(i+1);
        kari=spreadsheet.getRange(i,6);
        kengen=kari.setValue("look")
        

      }
      Logger.log("利用日"+useday);

      

    }
    var data=spreadsheet.getRange(2,1,lastRow,6);
    data.sort({column:1,ascending:false});
    return;
}




  
