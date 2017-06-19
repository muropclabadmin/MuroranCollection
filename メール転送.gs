'use strict'
var sendmailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sendmail'); //sheets[0];
var start = 0;
var max = 3
var threads = GmailApp.search('label:webアプリ-google-googlealerts-室蘭 is:unread', start, max);

function main()
{
  LogObj.clear();
  
  // シートを一旦クリア
  //sendmailSheet.clear();
  
  // 古い上位法の削除
  var lastRowCount = sendmailSheet.getLastRow();
  if(lastRowCount > 200)
  {
    var overSize = lastRowCount - 200;
    sendmailSheet.deleteRows(1, overSize);
  }
  
  sendCount = 0;
  execute(threads);
  
  GmailApp.markThreadsRead(threads);
  //GmailApp.moveThreadsToTrash(threads);
}

function execute(mailThread)
{
  for(var n=0; n<threads.length; n++){
    var the = threads[n];
    
    var msgs = the.getMessages();
    for(var m=0; m<msgs.length; m++){
      var msg = msgs[m];
      if(!msg.isUnread()) {
        continue;
      }
      
      var mailMessage = msg.getBody();
      var mailData = getMailData(mailMessage);
      
      for(var mail_count=0; mail_count<mailData.length; mail_count++)
      {
        var date = mailData[mail_count]['date'];
        var url = mailData[mail_count]['url'];
        var url_text = mailData[mail_count]['url_text'];
        var text = mailData[mail_count]['text'];
        var site_name = mailData[mail_count]['site_name'];
        
        var mailBody = sendWordpress(date, url, url_text, site_name, text); // メール送信
        sendmailSheet.appendRow([new Date(), mailBody]);
      }
    }
  }
}

function getMailData(text)
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
    LogObj.log(t);
    var td = t.match(/<td.*?>(.*?)<\/td>/gim);
    
    tempRegExp = new RegExp("<a style=\"color:#aaa;text-decoration:none\">(.+?)</a>", "gim");
    var tempDate = tempRegExp.exec(t);
    if(tempDate != null)
    {
      LogObj.log(tempDate);
      date = tempDate[1];
    }
    
    for(var j=0; j<td.length; j++)
    {
      var content = td[j].replace(/<td.*?>(.*?)<\/td>/gim, "$1");
      
      // URLの取り出し
      var url = content.match(/url=(.+?)&amp;.+/);
      if(url != null)
      {
        mailData[index] = {};
        mailData[index]['url'] = url[1];
        
        // リンクアドレス
        tempRegExp = new RegExp("<a .+?https://www.google.com/url\?.+?>(.+?)</a>", "gim");
        mailData[index]['url_text'] = removeHtmlTag(tempRegExp.exec(content)[1]).trim();
        
        // リンクする文字列
        tempRegExp = new RegExp("</div> <div style=.+?>(.+?)</div> </div>", "gim");
        mailData[index]['text'] = removeHtmlTag(tempRegExp.exec(content)[1]).trim();
        
        // <a style=\"text-decoration:none;color:#737373\"> <span>Yahoo!ロコ - Yahoo! JAPAN</span> </a>
        // サイト名
        tempRegExp = new RegExp("<a style=\"text-decoration:none;color:#737373\">(.+?)</a>", "gim");
        mailData[index]['site_name'] = removeHtmlTag(tempRegExp.exec(content)[1]).trim();
        
        // 日付
        mailData[index]['date'] = date;
        
        index++;
      }
    }
  }
        
  LogObj.log(mailData);
  return mailData;
}

/*
 * メールの内容からHTMLタグを取り除く
 */
function removeHtmlTag(html)
{
  var plainText = html.replace(/<("[^"]*"|'[^']*'|[^'">])*>/gim, ''); // その他のHTMLタグを全て削除
  if(plainText == null)
  {
    plainText = "";
  }
  return plainText;
}

/*
 * HTMLのリンク情報からリンクURLとテキストを取得する
 * [0]: 全体
 * [1]: URL
 * [2]: テキスト
 */
function getLinkData(html)
{
  var linkData = html.match(/<a href=\"(.*?)\".*?>(.*?)<\/a>/mi);
  if(linkData == null)
  {
    linkData = ["","",""];
  }
  return linkData;
}

/*
 * Wordpressに投稿する記事をメール送信する
 */
function sendWordpress(date, url, url_text, site_name, text)
{
  var mail = "tafi954roti@post.wordpress.com"; // Wordpress「室蘭情報まとめ」 http://blogmatome.wpblog.jp/
  var year = date.replace(/年.+?月.+?日/gi, "年");
  var month = date.replace(/月.+?日/gi, "月");
  var category = year+","+month+","+date;
  var publicize = "twitter";
  var thumnail = "https://blinky.nemui.org/shot/large?"+url;
  var tweet = date+" - "+url_text+"\n"+thumnail;
  var link = date+" - "+"<a href=\""+url+"\" target=\"_blank\">"+url_text + "\n"+thumnail+"</a>";
  var body = "[title "+url_text+"][category "+category+"][publicize "+publicize+"]"+tweet+"[/publicize]\n"+link+" - "+site_name+"\n\n"+text;
  var option = {htmlBody:body};
  
  GmailApp.sendEmail(mail, tweet, body, option); // sendEmail(recipient, subject, body, options)
  Utilities.sleep(1000);
  
  return body;
}

// ログ用シートにログ出力するためのオブジェクト
var LogObj = {
  logSheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log'),
  log: function(msg)
  {
    var tmpMsg = msg;
    if(typeof tmpMsg == typeof [])
    {
      tmpMsg = {log: tmpMsg};
    }
    var j = JSON.stringify(tmpMsg);
    this.logSheet.appendRow([new Date() , j]);
  },
  clear: function()
  {
    this.logSheet.clear();
  }
};
