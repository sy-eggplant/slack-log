// Configuration: Obtain Slack web API token at https://api.slack.com/web
var API_TOKEN = PropertiesService.getScriptProperties().getProperty('slack_api_token');
var FOLDER_NAME = 'Slack Logs';

// メール送信先
var mail_to = [
  "yoshikawa.shoko@slogan.jp",
  "kotaki.daisuke@slogan.jp",
  "yakubo.nanase@slogan.jp",
  "sugimura.yuri@slogan.jp",
  "kuwahara.shohei@slogan.jp",
  "yoshida.masashi@slogan.jp",
  "nakazato.kenichi@slogan.jp",
  "sugimoto@slogan.jp",
  "iwata.chiho@slogan.jp",
  "nihira@slogan.jp"
];

var dObj = new Date();
dObj.setDate(0);
var sheetName = "";
var body=sheetName+"のログ\n\n";

var subject = "【slacklog】";
  
function StoreLogsDelta() {
    var logger = new SlackChannelHistoryLogger();
    logger.run();
}
;
var SlackChannelHistoryLogger = (function () {
    function SlackChannelHistoryLogger() {
        this.memberNames = {};
    }
    // slackの情報をとってくる
    SlackChannelHistoryLogger.prototype.requestSlackAPI = function (path, params) {
        if (params === void 0) { params = {}; }
        var url = "https://slack.com/api/" + path + "?";
        var qparams = [("token=" + encodeURIComponent(API_TOKEN))];
        for (var k in params) {
            qparams.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
        }
        url += qparams.join('&');
        Logger.log("==> GET " + url);
        var resp = UrlFetchApp.fetch(url);
        var data = JSON.parse(resp.getContentText());
        if (data.error) {
            throw "GET " + path + ": " + data.error;
        }
        return data;
    };
    SlackChannelHistoryLogger.prototype.run = function () {
        var _this = this;
        var usersResp = this.requestSlackAPI('users.list');
        usersResp.members.forEach(function (member) {
            _this.memberNames[member.id] = member.name;
        });
        var teamInfoResp = this.requestSlackAPI('team.info');
        this.teamName = teamInfoResp.team.name;
        this.importChannelHistoryDelta();
    };
    // ログのフォルダ関連
    SlackChannelHistoryLogger.prototype.getLogsFolder = function () {
        var folder = DriveApp.getRootFolder();
        var path = [FOLDER_NAME, this.teamName];
        path.forEach(function (name) {
            var it = folder.getFoldersByName(name);
            if (it.hasNext()) {
                folder = it.next();
            }
            else {
                folder = folder.createFolder(name);
            }
        });
        return folder;
    };
    SlackChannelHistoryLogger.prototype.getSheet = function (d, readonly) {
        if (readonly === void 0) { readonly = false; }
        var dateString;
        if (d instanceof Date) {
            dateString = this.formatDate(d);
        }
        else {
            dateString = '' + d;
        }
        //シートの取得
        var spreadsheet;
        var sheetByID = {};
        var spreadsheetName = dateString;
        var folder = this.getLogsFolder();
        var it = folder.getFilesByName(spreadsheetName);
        if (it.hasNext()) {
            var file = it.next();
            spreadsheet = SpreadsheetApp.openById(file.getId());    
            Logger.log(file.getId());
        }

        var sheets = spreadsheet.getSheets();
        
        for ( var i in sheets ){
          if ( sheets[i].getSheetName() == sheetName ) {
            var sheet = sheets[i];
            
            var startrow = 1;
            var startcol = 1;
            var lastrow = sheet.getLastRow();
            var lastcol = sheet.getLastColumn();
            var sheetdata = sheet.getSheetValues(startrow, startcol, lastrow, lastcol);
            for (var k=0; k<lastrow; k++){
              for (var j=0; j<3; j++){
                if(j==0){
                  date=sheetdata[k][j];
                  
                  body += date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate() + '(' + date.getHours()+ ':' + date.getMinutes()+':'+ date.getSeconds()+')';
                 
                }
                else{  
                  body+=" | ";
                  body+=(sheetdata[k][j]);
                }
               
              }
              body+='\n';
            }
          }
        }
     
      MailApp.sendEmail(
        mail_to,
        subject,
        body
      );

    };
    SlackChannelHistoryLogger.prototype.importChannelHistoryDelta = function () {
        var _this = this;
        var now = new Date();
      // 指定した日付のログををとりたいとき用
      var year = now.getYear();
      var month = now.getMonth()-1;
      
      var now = new Date(year, month, 1);
      subject += (month+1) +"月分";
       Logger.log(now);
        var existingSheet = this.getSheet(now, true);
    };
    SlackChannelHistoryLogger.prototype.formatDate = function (dt) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM');
    };
    return SlackChannelHistoryLogger;
})();
