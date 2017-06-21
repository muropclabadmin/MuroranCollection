'use strict';

/**
* ログ用シートにログ出力するためのオブジェクト
*/
var MyLog = (function() {
  /** @constructor */
  var MyLog = function(logSheet) {
    if(!(this instanceof MyLog) || logSheet == null) {
      return new MyLog(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log'));
    }
    
    /**
    * ログ出力用シート
    * @type {Sheet}
    */
    this.logSheet = logSheet;
  };
  
  var p = MyLog.prototype;
  
  /**
  * 指定したログシートをクリアする
  */
  p.clear = function() { this.logSheet.clear(); };
  
  /**
  * 指定したログシートにデバッグログを出力する
  * @param {string|object} msg 出力するメッセージ(JSON形式のオブジェクト)
  */
  p.debug = function(msg) { this.log("debug", msg); };
  
  /**
  * 指定したログシートに情報ログを出力する
  * @param {string|object} msg 出力するメッセージ(JSON形式のオブジェクト)
  */
  p.info = function(msg) { this.log("info", msg); };
  
  /**
  * 指定したログシートに警告ログを出力する
  * @param {string|object} msg 出力するメッセージ(JSON形式のオブジェクト)
  */
  p.warning = function(msg) { this.log("warning", msg); };
  
  /**
  * 指定したログシートにエラーログを出力する
  * @param {string|object} msg 出力するメッセージ(JSON形式のオブジェクト)
  */
  p.error = function(msg) { this.log("error", msg); };
  
  /**
  * 指定したログシートにログを出力する
  * @param {ログの種類} type debug, info, warning, error
  * @param {string|object} msg 出力するメッセージ(JSON形式のオブジェクト)
  */
  p.log = function(type, msg)
  {
    var tmpMsg = msg;
    if(typeof tmpMsg == typeof [])
    {
      tmpMsg = {log: tmpMsg};
    }
    this.logSheet.appendRow([new Date() , type, JSON.stringify(tmpMsg)]);
  };
  
  return MyLog;
})();
