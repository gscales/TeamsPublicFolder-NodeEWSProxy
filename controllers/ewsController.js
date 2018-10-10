exports.getFolderItems = function (req, res) {
    (async () => {
        var ReturnItems = {};
        try {
            var ReturnItems = {};
            var Module = require('require-all')(__dirname + '/ews')
            FolderRequest = req.body;
            console.log(req.body);
            var token = req.headers.authorization.replace('Bearer ', '');            
            var EWSMAPI = require('ews-javascript-api');
            var service = new EWSMAPI.ExchangeService(EWSMAPI.ExchangeVersion.Exchange2013_SP1);
            service.EWSMAPI = EWSMAPI;
            service.Credentials = new EWSMAPI.OAuthCredentials(token);
            service.HttpHeaders = {};
            service.Url = new EWSMAPI.Uri("https://outlook.office365.com/Ews/Exchange.asmx");
            var FolderId = null;
            if (!FolderRequest.hasFolderId) {
                var PublicFolder = await Module.AutoDiscover.AutodiscoverUsersPublicFolder(service, FolderRequest.Mailbox);
                service.HttpHeaders["X-AnchorMailbox"] = PublicFolder;
                var pox = await Module.AutoDiscover.PoxAutoDiscover(PublicFolder, token);
                var ServerVal = await Module.Utils.ParseServerFromResult(pox);
                service.HttpHeaders["X-PublicFolderMailbox"] = ServerVal.Autodiscover.Response[0].Account[0].Protocol[0].Server[0];
                var Folder = await Module.PublicFolder.GetFolderFromPath(service, FolderRequest.FolderPath);
                var PR_REPLICA_LIST = new service.EWSMAPI.ExtendedPropertyDefinition(0x6698, service.EWSMAPI.MapiPropertyType.Binary);
                var PropVal = {};
                Folder.TryGetProperty(PR_REPLICA_LIST, PropVal);
                var ContentReplica = Buffer(PropVal.outValue).toString('ascii');
                var ContentHeader = ContentReplica.substring(0, ContentReplica.length - 1) + "@" + ServerVal.Autodiscover.Response[0].Account[0].Protocol[0].Server[0].split('@').pop()
                var poxContent = await Module.AutoDiscover.PoxAutoDiscover(ContentHeader, token);
                var PublicfolderVal = await Module.Utils.ParseServerFromResult(poxContent);
                service.HttpHeaders["X-AnchorMailbox"] = PublicfolderVal.Autodiscover.Response[0].User[0].AutoDiscoverSMTPAddress[0];
                service.HttpHeaders["X-PublicFolderMailbox"] = PublicfolderVal.Autodiscover.Response[0].User[0].AutoDiscoverSMTPAddress[0];
                ReturnItems.RoutingHeader = PublicfolderVal.Autodiscover.Response[0].User[0].AutoDiscoverSMTPAddress[0];
                FolderId = Folder.Id;
            } else {
                service.HttpHeaders["X-AnchorMailbox"] = FolderRequest.RoutingHeader;
                service.HttpHeaders["X-PublicFolderMailbox"] = FolderRequest.RoutingHeader;
                ReturnItems.RoutingHeader = FolderRequest.RoutingHeader;
                FolderId = new service.EWSMAPI.FolderId();
                FolderId.UniqueId = FolderRequest.UniqueId;
            }
            console.log("Discovery Complete");
            ReturnItems.FolderId = FolderId;
            ReturnItems.items = [];            
            var FolderItems = await Module.PublicFolderItems.FindItems(service, FolderId, FolderRequest.Offset, FolderRequest.PageCount,FolderRequest.Query);
            var iindex = 0;
            for (iindex = 0; iindex < FolderItems.items.length; ++iindex) {
                var pfItem = {};               
                pfItem.DateTimeReceived = Date.parse(FolderItems.items[iindex].DateTimeReceived.toString());
                pfItem.Subject = FolderItems.items[iindex].Subject;
                pfItem.SenderName = FolderItems.items[iindex].Sender.name;
                pfItem.SenderAddress = FolderItems.items[iindex].Sender.address;
                pfItem.Size = FolderItems.items[iindex].Size;
                var parsedPreview = "";
                if(FolderItems.items[iindex].Preview != null){
                    var lines = FolderItems.items[iindex].Preview.split('\n');
                    for(var i = 0;i < lines.length;i++){
                        if(i<3){
                            parsedPreview += (lines[i] + "\n");
                        }                    
                    }
                }                          

                pfItem.Preview  = parsedPreview; 
                pfItem.HasAttachments = FolderItems.items[iindex].HasAttachments;
                pfItem.WebClientReadFormQueryString = FolderItems.items[iindex].WebClientReadFormQueryString;
                ReturnItems.items.push(pfItem);
            }
            ReturnItems.moreAvailable = FolderItems.moreAvailable;
            ReturnItems.nextPageOffset = FolderItems.nextPageOffset;
            ReturnItems.Results = "Success";
        } catch (error) {
            console.log(error);
            ReturnItems.Results = "Error";
            ReturnItems.error = error;
        }
        res.setHeader('Content-Type', 'application/json');
        res.send(JSON.stringify(ReturnItems));
    })();
};
