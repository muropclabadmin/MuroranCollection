'use strict';

/**
* エントリポイント
*/
function main()
{
  var mc = new MuroranCollection();
  mc.main();
}

/**
* 処理用のクラス
*/
var MuroranCollection = (function() {
  /** @constructor */
  var MuroranCollection = function() {
    if(!(this instanceof MuroranCollection)) {
      return new MuroranCollection();
    }
    
    /**
    * ログ出力オブジェクト
    * @type {MyLog}
    */
    this.myLog = new MyLog(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log'));
    
    /**
    * メール送信した内容を記録するシート
    * @type {Sheet}
    */
    this.sendmailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sendmail');
    
    /**
    * メール送信情報を記録するシート
    * @type {Sheet}
    */
    this.statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('status');
    
    /**
    * メール送信した内容を記録するシートのメール最終送信日のセル
    * @type {Range}
    */
    this.mailDateCell = this.statusSheet.getRange(2, 1); // メールの最終送信日
    
    /**
    * メール送信した内容を記録するシートのメール送信回数のセル
    * @type {Range}
    */
    this.mailCountCell = this.statusSheet.getRange(2, 2); // メールの送信回数
  };
  
  var p = MuroranCollection.prototype;
  
  /**
  * メイン処理
  */
  p.main = function()
  {
    var start = 0;
    var max = 3
    var threads = GmailApp.search('label:webアプリ-google-googlealerts-室蘭 is:unread', start, max);
    
    this.myLog.clear();
    
    // メールの送信回数の確認
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
    var preSendDate = Utilities.formatDate(this.mailDateCell.getValue(), 'Asia/Tokyo', 'yyyy年M月d日');
    if(preSendDate != now)
    {
      // 日付が変わっていた場合、メール送信回数をリセットする
      this.mailCountCell.setValue(0);
    }
    
    // 古い情報の削除
    var lastRowCount = this.sendmailSheet.getLastRow();
    this.myLog.debug(lastRowCount);
    if(lastRowCount > 200)
    {
      var overSize = lastRowCount - 200;
      this.myLog.info(overSize);
      this.sendmailSheet.deleteRows(1, overSize);
    }
    
    this.execute(threads);
    
    //GmailApp.markThreadsRead(threads); // スレッドを既読にする
    //GmailApp.moveThreadsToTrash(threads); // スレッドをゴミ箱に移動
    
    return;
  };
  
  /**
  * メールスレッドに対して処理を行う。メッセージごとに既読にする
  * @param {Threads} mailThread メールのスレッド(複数)
  */
  p.execute = function(mailThread)
  {
    for(var n=0; n<mailThread.length; n++)
    {
      var thread = mailThread[n];
      
      var msgs = thread.getMessages();
      for(var m=0; m<msgs.length; m++){
        var msg = msgs[m];
        if(!msg.isUnread())
        {
          // 既読であれば無視する
          msg.moveToTrash(); // ゴミ箱に移動する
          continue;
        }
        
        // メッセージごとに処理をする
        var mailMessage = msg.getBody();
        var mailData = this.getMailData(mailMessage);
        
        for(var mail_count=0; mail_count<mailData.length; mail_count++)
        {
          var date = mailData[mail_count]['date'];
          var url = mailData[mail_count]['url'];
          var url_text = mailData[mail_count]['url_text'];
          var text = mailData[mail_count]['text'];
          var site_name = mailData[mail_count]['site_name'];
          
          var mailBody = this.sendWordpress(date, url, url_text, site_name, text); // メール送信
          
          // シートにメール情報を記録
          this.mailDateCell.setValue(new Date());
          this.mailCountCell.setValue(Number(this.mailCountCell.getValue()) + 1);
          this.sendmailSheet.appendRow([new Date(), mailBody]); // 送信メール本文をシートに記録
        }
        
        msg.markRead(); // メッセージごとに既読にする
        msg.moveToTrash(); // ゴミ箱に移動する
      }
    }
  };
  
  /**
  * メールの内容を取得する
  * @param {String} text メール本文
  * @return {Array<Object>} メール情報のオブジェクトを格納した配列
  */
  p.getMailData = function(text)
  {
    var plainText = text;
    var tempRegExp = null;
    
    plainText = plainText.replace(/\r\n|\r|\n/gim, "\n"); // 改行コードを削除
    plainText = plainText.match(/<div style=.+>(.+?)<\/div>/gi)[0]; // メインとなる部分を取り出す
    plainText = plainText.replace(/[\f\r\v\u00a0\u1680\u180e\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff]/gi, ''); // その他の制御コードを削除
    plainText = plainText.replace(/[\t]/gi, '  '); // TABを空白2文字に変換
    
    // 各行からURL情報を取り出す
    var linetext = plainText.split('<td style="padding-left:18px"></td>'); // 分解
    
    // TABLEタグの中からTDタグの中身を取り出す
    var trList = plainText.match(/<tr>(.+?)<\/tr>/gi);
    var mailData = [];
    var date = ""; // メールの日付
    var index = 0;
    for(var i=0; i<trList.length; i++)
    {
      var t = trList[i];
      this.myLog.debug(t);
      var td = t.match(/<td.*?>(.*?)<\/td>/gim);
      
      tempRegExp = new RegExp("<a style=\"color:#aaa;text-decoration:none\">(.+?)</a>", "gim");
      var tempDate = this.execRegExp(t, tempRegExp);
      if(tempDate != "")
      {
        this.myLog.debug(tempDate);
        date = tempDate;
      }
      
      for(var j=0; j<td.length; j++)
      {
        var content = td[j].replace(/<td.*?>(.*?)<\/td>/gim, "$1");
        
        // URL(リンクアドレス)の取り出し
        var url = content.match(/url=(.+?)&amp;.+/);
        if(url != null)
        {
          mailData[index] = {};
          var tmpResult = "";
          
          // URL(リンクアドレス)
          var tempUrl = url[1];
          if(tempUrl.indexOf('%') != -1)
          {
            // 「%」を含む場合、URLをデコードする
            var decUrl = decodeURI(url[1]);
            if(decUrl.match(/%3A|%2F|%3B|%3F/))
            {
              // :=%3A /=%2F ;=%3B ?=%3F
              decUrl = decodeURIComponent(decUrl);
            }
            tempUrl = decUrl;
          }
          mailData[index]['url'] = tempUrl;
          
          // リンクする文字列
          tempRegExp = new RegExp("<a .+?https://www.google.com/url\?.+?>(.+?)</a>", "gim");
          tmpResult = this.execRegExp(content, tempRegExp);
          mailData[index]['url_text'] = this.removeHtmlTag(tmpResult).trim();
          
          // 説明文
          tempRegExp = new RegExp("</div> <div style=.+?>(.+?)</div> </div>", "gim");
          tmpResult = this.execRegExp(content, tempRegExp);
          mailData[index]['text'] = this.removeHtmlTag(tmpResult).trim();
          
          // サイト名
          tempRegExp = new RegExp("<a style=\"text-decoration:none;color:#737373\">(.+?)</a>", "gim");
          tmpResult = this.execRegExp(content, tempRegExp);
          mailData[index]['site_name'] = this.removeHtmlTag(tmpResult).trim();
          
          // 日付
          mailData[index]['date'] = date;
          
          index++;
        }
      }
    }
    
    this.myLog.debug(mailData);
    return mailData;
  };
  
  /**
  * メールの内容からHTMLタグを取り除く
  * @param {String} html HTMLタグを含むメール本文
  * @return {String} プレーンテキスト
  */
  p.removeHtmlTag = function(html)
  {
    var plainText = html.replace(/<("[^"]*"|'[^']*'|[^'">])*>/gim, ''); // その他のHTMLタグを全て削除
    if(plainText == null)
    {
      plainText = "";
    }
    return plainText;
  };
  
  /**
  * HTMLのリンク情報からリンクURLとテキストを取得する
  * [0]: 全体
  * [1]: URL
  * [2]: テキスト
  * @param {String} html HTMLタグを含む文字列
  * @return {Array<String>} リンク情報の配列
  */
  p.getLinkData = function(html)
  {
    var linkData = html.match(/<a href=\"(.*?)\".*?>(.*?)<\/a>/mi);
    if(linkData == null)
    {
      linkData = ["","",""];
    }
    return linkData;
  };
  
  /**
  * 正規表現で文字列を取り出す。見つからない場合は空文字列を返す
  * @param {String} content HTMLタグを含む文字列
  * @param {RegExp} tempRegExp 正規表現オブジェクト
  * @return {String} 取り出した文字列
  */
  p.execRegExp = function(content, tempRegExp)
  {
    var resultStr = "";
    var tmp = tempRegExp.exec(content);
    if(tmp != null)
    {
      resultStr = tmp[1];
    }
    return resultStr;
  };
  
  /**
  * Wordpressに投稿する記事をメール送信する
  * @param {String} date 記事の日付情報
  * @param {String} url 元記事のリンクURL
  * @param {String} url_text 元記事のタイトル
  * @param {String} site_name 元記事のサイト名
  * @param {String} text 元記事の内容
  * @return {String} 送信したメール本文
  */
  p.sendWordpress = function(date, url, url_text, site_name, text)
  {
    var mail = "tafi954roti@post.wordpress.com"; // Wordpress「室蘭情報まとめ」 http://blogmatome.wpblog.jp/
    var year = date.replace(/年.+?月.+?日/gi, "年");
    var month = date.replace(/月.+?日/gi, "月");
    var category = year+","+month+","+date;
    var publicize = "twitter";
    var thumnail_url = "https://blinky.nemui.org/shot/large?"+url;
    var thumnail_link = "<a href='"+url+"' target='_blank'><img src='"+thumnail_url+"' /></a>";
    var tweet = date+" - "+url_text;
    var link = date+" - <a href='"+url+"' target='_blank'>"+url_text+"</a>";
    var jetpack_tag = "[title "+url_text+"][category "+category+"][publicize "+publicize+"]"+tweet+"[/publicize]";
    var body = jetpack_tag+"\n"+link+" - "+site_name+"<br />\n<br />\n"+thumnail_link+"<br />\n"+text;
    var option = {htmlBody:body};
    
    // メール送信
    GmailApp.sendEmail(mail, tweet, body, option); // sendEmail(recipient, subject, body, options)
    this.myLog.info({mail:mail, tweet:tweet, option:option});
    
    Utilities.sleep(1000); // 連続送信を避けるため、ちょっと待機
    
    return body;
  };
  
  return MuroranCollection;
})();
