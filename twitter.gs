// postTweet(message): Posts a Tweet given that all settings are correct.

function postTweet(message) {
  
    var method = "POST";
    var baseUrl = "https://api.twitter.com/1.1/statuses/update.json";
    var props = PropertiesService.getScriptProperties();
    
    var oauthToken = settingsArray[3][3]
    var oauthKey = settingsArray[3][1]
    var oauthSecretKey = settingsArray[3][2]
    var oauthSecretToken = settingsArray[3][4]
    
    var oauthParameters = {
      oauth_consumer_key: oauthKey,
      oauth_token: oauthToken,
      oauth_timestamp: (Math.floor((new Date()).getTime() / 1000)).toString(),
      oauth_signature_method: "HMAC-SHA1",
      oauth_version: "1.0"
    }
    
    oauthParameters.oauth_nonce = oauthParameters.oauth_timestamp + Math.floor(Math.random() * 100000000);
    
    var payload = {
      status: message
    }
    
    var queryKeys = Object.keys(oauthParameters).concat(Object.keys(payload)).sort();
    
    var baseString = queryKeys.reduce(function(acc, key, idx) {
      if (idx) acc += encodeURIComponent("&");
      if (oauthParameters.hasOwnProperty(key))
        acc += encode(key + "=" + oauthParameters[key]);
      else if (payload.hasOwnProperty(key))
        acc += encode(key + "=" + encode(payload[key]));
      return acc;
    }, method.toUpperCase() + '&' + encode(baseUrl) + '&')
    
    oauthParameters.oauth_signature = Utilities.base64Encode(
      Utilities.computeHmacSignature(
        Utilities.MacAlgorithm.HMAC_SHA_1,
        baseString,
        oauthSecretKey + "&" + oauthSecretToken
      )
    )
    
    var options = {
      method: method,
      headers: {
        authorization: "OAuth " + Object.keys(oauthParameters).sort().reduce(function(acc, key) {
        acc.push(key + '="' + encode(oauthParameters[key]) + '"');
        return acc;
      }, []).join(', ')
    },
        payload: Object.keys(payload).reduce(function(acc, key) {
          acc.push(key + '=' + encode(payload[key]));
          return acc;
        }, []).join('&'),
          muteHttpExceptions: true
  }
  
  var response = UrlFetchApp.fetch(baseUrl, options);
  var responseHeader = response.getHeaders();
  var responseText = response.getContentText();
  
  
  }
  
  function encode(string) {
    return encodeURIComponent(string)
      .replace('!', '%21')
      .replace('*', '%2A')
      .replace('(', '%28')
      .replace(')', '%29')
      .replace("'", '%27');
  } 
  