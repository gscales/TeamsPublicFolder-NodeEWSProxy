(function () {

    async function GetFolderFromPath(service,folderPath) {
        try {               
                var folderid = new service.EWSMAPI.FolderId(service.EWSMAPI.WellKnownFolderName.PublicFoldersRoot);   
                var RootFolder = await service.EWSMAPI.Folder.Bind(service,folderid).then(function (response) {
                    return response;
                });
                var folderPaths = folderPath.split("\\");
                var i=0;                
                for(i=1; i < folderPaths.length;++i){
                    //console.log(folderPaths[i]);
                    RootFolder = await FindFolder(service,RootFolder,folderPaths[i]);
                    
                }
                return RootFolder;
       
        }
        catch (error) {
            (error.message);
            return null;
        }
    }

    async function FindFolder(service,Folder,folderName){
        console.log(Folder.DisplayName);
        var FolderView = new service.EWSMAPI.FolderView(1);
        var PR_REPLICA_LIST = new service.EWSMAPI.ExtendedPropertyDefinition(0x6698, service.EWSMAPI.MapiPropertyType.Binary);
        var psPropset= new service.EWSMAPI.PropertySet(service.EWSMAPI.BasePropertySet.FirstClassProperties);
        psPropset.Add(PR_REPLICA_LIST);
        FolderView.PropertySet = psPropset;
        var Searchfilter = new service.EWSMAPI.SearchFilter.IsEqualTo(service.EWSMAPI.FolderSchema.DisplayName,folderName);
        var TargetFolder = await service.FindFolders(Folder.Id,Searchfilter,FolderView).then(function (response){
            return response.Folders[0];
        });
        return TargetFolder;
    }

    module.exports.GetFolderFromPath = function (service,Mailbox) {
        return GetFolderFromPath(service,Mailbox);
    }

}());

