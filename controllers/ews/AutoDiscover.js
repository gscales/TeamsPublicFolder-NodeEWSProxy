(function () {

    const PoxAutoDiscoverRequest = (emailAddress,Token) => {
        return new Promise(
            (resolve, reject) => {
                var PoxRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                PoxRequest += "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006\">";
                PoxRequest += "  <Request>";
                PoxRequest += "    <EMailAddress>" + emailAddress + "</EMailAddress>";
                PoxRequest += "    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>";
                PoxRequest += "  </Request>";
                PoxRequest += "</Autodiscover>";
                var request = require('request');
                var options = {
                    url: 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml',
                    headers: {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134',
                        'Content-Type': 'text/xml',
                    },
                    auth: {
                        'bearer': Token
                    },
                    method: "POST",
                    body: PoxRequest,
                    
                };
                request(options, function (err, res) {
                    if(err == null){                                          
                        resolve(res.body);
                    }else{
                        reject(err);
                    }
                }
                );
				
            }
        );
    }

    async function AutodiscoverUsersPublicFolder(service,mailbox) {
        try {  
            console.log("AutodiscoverUsersPublicFolder");             
            var AutoDiscoverService = new service.EWSMAPI.AutodiscoverService(service.EWSMAPI.ExchangeVersion.Exchange2013_SP1);
            AutoDiscoverService.Credentials = service.Credentials;
            AutoDiscoverService.EnableScpLookup = false;
            AutoDiscoverService.RedirectionUrlValidationCallback = function (url){
                return true;
                }
            var settings = [
            service.EWSMAPI.UserSettingName.InternalEwsUrl,
            service.EWSMAPI.UserSettingName.ExternalEwsUrl,
            service.EWSMAPI.UserSettingName.PublicFolderInformation,
            ];
            var PublicFolderInformation = await AutoDiscoverService.GetUserSettings([mailbox], settings).then(function (response) {
            let pfReturn = "";
            for (var _i = 0, _a = response.Responses; _i < _a.length; _i++) {
                var resp = _a[_i];
                pfReturn = resp.Settings[92]; 
            }
            return pfReturn;
        }, function (e) {
            console.log(e);
            //log errors or do something with errors
        });
        return PublicFolderInformation;
        }
        catch (error) {
            console.log(error);
            (error.message);
            return null;
        }
    }

    async function PoxAutoDiscover(emailAddress,Token) {
        try {           
            let PoxResult = await PoxAutoDiscoverRequest(emailAddress,Token);
            return PoxResult;
        }
        catch (error) {
            console.log(error.message);
        }
    }

    module.exports.AutodiscoverUsersPublicFolder = function (service,Mailbox) {
        return AutodiscoverUsersPublicFolder(service,Mailbox);
    }


    module.exports.PoxAutoDiscover = function (emailAddress,Token) {
        return PoxAutoDiscover(emailAddress,Token);
    }
   
}());

