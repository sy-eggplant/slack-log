var API_TOKEN = "";
// Slack offers 10,000 history logs for free plan teams
var MAX_HISTORY_PAGINATION = 10;
var HISTORY_COUNT_PER_PAGE = 1000;
var stamps = {};
var counts = [];
var names = [];
var sheet = SpreadsheetApp.getActiveSheet();
// 時間の判定
var timezone = sheet.getParent().getSpreadsheetTimeZone();
var now = new Date;
var yestarday_date = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
var yestarday = Utilities.formatDate(yestarday_date, timezone, 'yyyy-MM-dd');

function StoreLogsDelta() {
  var logger = new SlackChannelHistoryLogger();
  logger.run();
  //シートを削除
  sheet.clear();
  sheet.getRange(1,1).setValue(yestarday);
  
  var names = Object.keys(stamps);
  var ary = [];
  Logger.log(stamps);
  for (var i=0; i<names.length; i++) {
    ary.push([names[i]]);
  }
  sheet.getRange(2,1,ary.length,1).setValues(ary);
  for (var i=0; i<names.length; i++) {
    counts.push([stamps[names[i]]]); 
  }
  sheet.getRange(2,2,counts.length,1).setValues(counts);
};

var SlackChannelHistoryLogger = (function () {
    function SlackChannelHistoryLogger() {
        this.memberNames = {};
    }
    SlackChannelHistoryLogger.prototype.requestSlackAPI = function (path, params) {
        if (params === void 0) { params = {}; }
        var url = "https://slack.com/api/" + path + "?";
        var qparams = [("token=" + encodeURIComponent(API_TOKEN))];
        for (var k in params) {
            qparams.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
        }
        url += qparams.join('&');
        try{
          var resp = UrlFetchApp.fetch(url);
          var data = JSON.parse(resp.getContentText());
          if (data.error) {
            throw "GET " + path + ": " + data.error;
          }
          return data;
         }catch(e){
          return "err";
          }
    };
    SlackChannelHistoryLogger.prototype.run = function () {
        var _this = this;
        var channelsResp = this.requestSlackAPI('channels.list');
            for (var _i = 0, _a = channelsResp.channels; _i < _a.length; _i++) {
              var ch = _a[_i];
              this.importChannelHistoryDelta(ch);
            }
    };  
    SlackChannelHistoryLogger.prototype.importChannelHistoryDelta = function (ch) {
        var _this = this;
        var now = new Date();
        var oldest = '1'; // oldest=0 does not work
        var messages = this.loadMessagesBulk(ch, { oldest: oldest });
        var dateStringToMessages = {};
      Logger.log(ch);
      
      
      if(messages != "err"){
        messages.forEach(function (msg) {
          var date = new Date(+msg.ts * 1000);
          var rec = msg.reactions ? msg.reactions : "";
          var m_date = Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
              if(rec !== "" && m_date == yestarday){
                var name = rec[0].name;
          Logger.log(m_date);
                var tmp = 0;
                if (stamps[name]) {
                  tmp = stamps[name];
                  stamps[name] = tmp + rec[0].count;
                } else {
                  stamps[name] = rec[0].count;
                }      
              }
        });
      }
     
    };
    SlackChannelHistoryLogger.prototype.formatDate = function (dt) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM');
    };
    SlackChannelHistoryLogger.prototype.loadMessagesBulk = function (ch, options) {
        var _this = this;
        if (options === void 0) { options = {}; }
        var messages = [];
        options['count'] = HISTORY_COUNT_PER_PAGE;
        options['channel'] = ch.id;
        var loadSince = function (oldest) {
            if (oldest) {
                options['oldest'] = oldest;
            }
            var resp = _this.requestSlackAPI('channels.history', options);
          if(resp != "err"){
            messages = resp.messages.concat(messages);
          }
            return resp;
        };
        var resp = loadSince();
        var page = 1;
        while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
            resp = loadSince(resp.messages[0].ts);
            page++;
        }
        return messages.reverse();
    };
    return SlackChannelHistoryLogger;
})();


function Ranking() {
  var range = sheet.getRange(1, 1, sheet.getLastRow(), 2);
  var s_cnt = range.getValues();
  
  var tmp = 0;
  var tmp_n = "";
 
  for (var j=1; j<s_cnt.length; j++) {
    for (var k=j+1; k<s_cnt.length; k++) {
      if(s_cnt[j][1] < s_cnt[k][1]){
         tmp =  s_cnt[j][1];
        s_cnt[j][1] = s_cnt[k][1];
        s_cnt[k][1] = tmp;
         tmp_n =  s_cnt[j][0];
        s_cnt[j][0] = s_cnt[k][0];
        s_cnt[k][0] = tmp_n;
      }
    }
  }
  sheet.getRange(1,1,s_cnt.length,2).setValues(s_cnt);
  var snt = s_cnt[0] + "\n1: :"+s_cnt[1][0] + ": (" +s_cnt[1][1] + ") \n" +
    "2: :"+s_cnt[2][0] + ": (" +s_cnt[2][1] + ") \n" +
      "3: :"+s_cnt[3][0] + ": (" +s_cnt[3][1] + ") \n" +
        "4: :"+s_cnt[4][0] + ": (" +s_cnt[4][1] + ") \n" +
          "5: :"+s_cnt[5][0] + ": (" +s_cnt[5][1] + ") \n" 
    ;
  Logger.log(snt);
 slack(snt,s_cnt[1][0]);
}

function slack(message,name) {
 
    var url        = 'https://slack.com/api/chat.postMessage';
    var token      = API_TOKEN;
    var channel    = '#times_yosikawa';
    var text       = message;
    
   
      var username   = '昨日使われたスタンプランキング';
      var icon_emoji = ':'+name+':';
   
    var parse      = 'full';
    var method     = 'post';
 
    var payload = {
        'token'      : token,
        'channel'    : channel,
        'text'       : text,
        'username'   : username,
        'parse'      : parse,
        'icon_emoji' : icon_emoji,
        'link_names' : true
    };
 
    var params = {
        'method' : method,
        'payload' : payload
    };
    var response = UrlFetchApp.fetch(url, params);
}
