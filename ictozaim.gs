const CONSUMER_KEY = PropertiesService.getScriptProperties().getProperty("ZAIM_CONSUMER_ID");
const CONSUMER_SECRET = PropertiesService.getScriptProperties().getProperty("ZAIM_CONSUMER_SECRET")
const MOBIRY_PAYMENT_ID = parseFloat(PropertiesService.getScriptProperties().getProperty("MOBIRY_PAYMENT_ID"));
const DEFAULT_CATEGORY_ID = parseFloat(PropertiesService.getScriptProperties().getProperty("DEFAULT_CATEGORY_ID"));
const DEFAULT_GENRE_ID = parseFloat(PropertiesService.getScriptProperties().getProperty("DEFAULT_GENRE_ID"));
const manae_wallet="18405154";

// 今日の日付を取得
var date = new Date(); // 現在の日付と時刻を取得
var today = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
var time =Utilities.formatDate(date,Session.getScriptTimeZone(),"HH:mm");
Logger.log("time:"+time);

// 1日前の日付を取得
var yesterday = new Date(date.getFullYear(), date.getMonth(), date.getDate() - 70);
yesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");
Logger.log("yesterday:"+ yesterday);
Logger.log("today:"+ today);

// 認証のリセット
function reset() {
  var service = getService();
  service.reset();
}

// 認証サービスの設定
function getService() {
  return OAuth1.createService("Zaim")
    // Set the endpoint URLs.
    .setAccessTokenUrl("https://api.zaim.net/v2/auth/access")
    .setRequestTokenUrl("https://api.zaim.net/v2/auth/request")
    .setAuthorizationUrl("https://auth.zaim.net/users/auth")

    // Set the consumer key and secret.
    .setConsumerKey(CONSUMER_KEY)
    .setConsumerSecret(CONSUMER_SECRET)

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction("authCallback")

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties());
}

// OAuth Callbackの設定
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput("認証できました！このページを閉じて再びスクリプトを実行してください。");
  } else {
    return HtmlService.createHtmlOutput("認証に失敗");
  }
}

// GETパラメーターを作成
function encodeParams(params) {
  var encodedParams = [];
  for (var key in params) {
    encodedParams.push(encodeURIComponent(key) + "=" + encodeURIComponent(params[key]));
  }
  return encodedParams.join("&");
}

// 過去の支払いデータを取得
function getPastData(service) {
  var url = "https://api.zaim.net/v2/home/money";

  // 日付で検索
  var params = {
    "start_date": "2025-02-04",
    "end_date": today,
  }
     

  // データの取得
  var response = service.fetch(url + "?" + encodeParams(params));
  var result = JSON.parse(response.getContentText());
  // Logger.log(result); // 取得したデータを見たい場合コメントアウトを外す

  // 楽天ペイのみの支払いでフィルタリング
  var rakutenPayData = result.money.filter(function(item) {
    return item.from_account_id === MOBIRY_PAYMENT_ID || manae_wallet;
  });

  return rakutenPayData
}

// 楽天ペイ情報をZaimに登録（メイン関数）
function rakutenPayToZaim() {
  seikikamobi();
  // Gmailで該当メールを検索
  var start = 0;
  var max = 5; // 過去いくつまでメールを遡るか
  var query = 'subject: ("楽天ペイアプリご利用内容確認メール" OR "楽天ペイ 注文受付") ';
  var threads = GmailApp.search(query, start, max);

  // 既存のデータを取得
  var service = getService();
  if (service.hasAccess()) {
    var existingData = getPastData(service)
  } else {
    var authorizationUrl = service.authorize();
    Logger.log("次のURLを開いてZaimで認証したあと、再度スクリプトを実行してください。: %s",
      authorizationUrl);
  }
  // スプレッドシートの最下部を取得
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();

  // 各楽天ペイの支払いについて登録
  for (var i = 2; i < lastRow+1; i++){
      var usageDate=sheet.getRange(i,1).getValue();
      var usageDate = Utilities.formatDate(usageDate, 'JST', 'yyyy-MM-dd');
      var comment=sheet.getRange(i,2).getValue();
      var amount=sheet.getRange(i,3).getValue();
      var shop=sheet.getRange(i,5).getValue();
      var shubetu=sheet.getRange(i,4).getValue();
      if (shubetu=="乗降"|| shubetu=="乗降乗越精算"){
        var shurui="paymentda";
  }
else{
  var shurui="autocharge";
  continue;
}
      console.log("「今回」　　日付:"+usageDate+",支払金額:"+amount+",お店:"+shop+"メモ:"+comment);


      var isExisting = false;

      if (comment=="物販")
        {
         shop="物販";
         comment="システムから入力";

          var category_id =102;
          var genre_id = 10201;

        }
      else if(comment=="バスチャージ-車載端末")
      {
        shop="バス　車両端末"

        var category_id =199;
        var genre_id = 19908;
      }


      // 既存のデータと比較して、新しいデータのみを追加
      for (var k = 0; k < existingData.length; k++) {
        
        if (existingData[k]["date"] == usageDate  && existingData[k]["comment"] == comment) {
          isExisting = true;
                    Logger.log(existingData[k]["comment"]+":"+comment);
          Logger.log("既に入力済み")
          break;
        }
      }

      if (!isExisting) {
        // 登録カテゴリーシートからデータを取得
        var files = DriveApp.getFilesByName("ZAIM_DB");
        var originalData = []
        if (files.hasNext()) {
          spreadsheet = SpreadsheetApp.open(files.next());
          var originalSheet = spreadsheet.getSheetByName("登録カテゴリー");
          originalData = originalSheet.getRange(2, 1, originalSheet.getLastRow() - 1, 4).getValues();
        }

        // 登録カテゴリーシートを検索して適切なカテゴリとジャンルを設定
        var category_id = DEFAULT_CATEGORY_ID;
        var genre_id = DEFAULT_GENRE_ID;
        for (var k = 0; k < originalData.length; k++) {
          var [storeName, exactMatch, categoryId, genreId] = originalData[k];
          if ((exactMatch && storeName === shop) || (!exactMatch && shop.includes(storeName))) {
            category_id = categoryId;
            genre_id = genreId;
            break;
          }
        }

        if (comment=="物販")
        {
         shop="物販";
         comment="システムから入力";

          var category_id =102;
          var genre_id = 10201;

        }
      else if(comment=="バスチャージ-車載端末")
      {
        shop="バス　車両端末"

        var category_id =199;
        var genre_id = 19908;
      }

      if(shurui=="paymentda"){



        // 支払い情報の登録
        var url = "https://api.zaim.net/v2/home/money/payment";
        var payload = {
          "category_id": category_id,
          "genre_id": genre_id,
          "amount": amount,
          "date": usageDate,
          "place": shop,
          "from_account_id": MOBIRY_PAYMENT_ID,
          "comment": comment
        };
      }else if(shurui=="transfer")
      {
         var url = "https://api.zaim.net/v2/home/money/transfer";
        var payload = {
          "mode":"transfer",    
          "amount": amount,
          "date": usageDate,
          "place": shop,
          "to_account_id":MOBIRY_PAYMENT_ID,
          "from_account_id": manae_wallet,
          "comment": "バスチャージ 車載端末",
          "currency_code":"JPY"

        };

      }
        var options = {
          "method": "post",
          "payload": payload
        };
        service.fetch(url, options);
        Logger.log("支払い入力完了")
      }
    }
}

// Spreadsheetに書き込み
function writeCategoriesToSpreadsheet(categories, genres, accounts) {
  const SPREADSHEET_NAME = "ZAIM_DB";
  const CATEGORY_SHEET_NAME = "カテゴリと内訳";
  const ACCOUNT_SHEET_NAME = "支払方法";
  const ORIGINAL_SHEET_NAME = "登録カテゴリー";

  // スプレッドシートを取得または作成
  let spreadsheet;
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  }

  // カテゴリとジャンルのシートを作成
  let categorySheet = spreadsheet.getSheetByName(CATEGORY_SHEET_NAME);
  if (categorySheet) {
    categorySheet.clear(); // 既存のシートをクリア
  } else {
    categorySheet = spreadsheet.insertSheet(CATEGORY_SHEET_NAME);
  }

  // アカウントのシートを作成
  let accountSheet = spreadsheet.getSheetByName(ACCOUNT_SHEET_NAME);
  if (accountSheet) {
    accountSheet.clear(); // 既存のシートをクリア
  } else {
    accountSheet = spreadsheet.insertSheet(ACCOUNT_SHEET_NAME);
  }

  // オリジナルカテゴリーシートの作成
  let originalSheet = spreadsheet.getSheetByName(ORIGINAL_SHEET_NAME);
  if (!originalSheet) {
    originalSheet = spreadsheet.insertSheet(ORIGINAL_SHEET_NAME);
    var originalHeaders = ["店舗名", "完全一致", "カテゴリーID", "内訳ID"];
    originalSheet.getRange(1, 1, 1, originalHeaders.length).setValues([originalHeaders]);
  }

  // 最初のシートが存在する場合、削除する
  let defaultSheet = spreadsheet.getSheetByName("シート1");
  if (defaultSheet) {
    spreadsheet.deleteSheet(defaultSheet);
  }

  // カテゴリとジャンルのヘッダーを作成
  var categoryHeaders = ["カテゴリーID", "カテゴリー名", "内訳ID", "内訳名"];
  categorySheet.appendRow(categoryHeaders);

  // アカウントのヘッダーを作成
  var accountHeaders = ["支払ID", "支払方法"];
  accountSheet.appendRow(accountHeaders);

  // データ行を作成
  var categoryRows = [];
  categories.forEach(category => {
    var categoryGenres = genres.filter(genre => genre.category_id === category.id);
    if (categoryGenres.length > 0) {
      categoryGenres.forEach(genre => {
        categoryRows.push([
          category.id,
          category.name,
          genre.id,
          genre.name,
        ]);
      });
    } else {
      categoryRows.push([
        category.id,
        category.name,
        "",
        "",
      ]);
    }
  });
  var accountRows = accounts.map(account => [account.id, account.name]);

  // 一括で書き込み
  categorySheet.getRange(2, 1, categoryRows.length, categoryHeaders.length).setValues(categoryRows);
  accountSheet.getRange(2, 1, accountRows.length, accountHeaders.length).setValues(accountRows);

  Logger.log("カテゴリ情報、内訳、支払い方法一覧を取得しました: " + spreadsheet.getUrl());
}

// カテゴリと内訳の取得（スプレッドシート作成のメイン関数）
function getInfo() {
  var service = getService();
  if (!service.hasAccess()) {
    var authorizationUrl = service.authorize();
    Logger.log("次のURLを開いてZaimで認証したあと、再度スクリプトを実行してください。: %s",
      authorizationUrl);
  }

  // カテゴリの取得
  var categoryUrl = "https://api.zaim.net/v2/home/category";
  var response = service.fetch(categoryUrl)
  var category = JSON.parse(response.getContentText()).categories;
  console.log("---------------------------------カテゴリー------------------------------------------------------------")
  console.log(category)
  var filterCategory = category.filter(item => item.mode === "payment" && item.active === 1.0)
    .sort((a, b) => a.sort - b.sort);

  // ジャンルの取得
  var genreUrl = "https://api.zaim.net/v2/home/genre";
  response = service.fetch(genreUrl)
  var genre = JSON.parse(response.getContentText()).genres;
  console.log("-----------------------------------ジャンル------------------------------------------------------------")
  console.log(genre)
  var filterGenre = genre.filter(item => item.active === 1.0)
    .sort((a, b) => a.sort - b.sort);
    

  // アカウント一覧の取得
  var accountUrl = "https://api.zaim.net/v2/home/account";
  var response = service.fetch(accountUrl);
  var account = JSON.parse(response.getContentText()).accounts;
  account = account.filter(item => item.active === 1.0)
    .sort((a, b) => a.sort - b.sort);

  // スプレッドシートに書き込み
  writeCategoriesToSpreadsheet(filterCategory, filterGenre, account);
}