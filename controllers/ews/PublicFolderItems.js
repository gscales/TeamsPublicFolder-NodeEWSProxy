(function () {

    async function FindItems(service,FolderId,offset,totalCount,Query){
       
        var ItemView = new service.EWSMAPI.ItemView(totalCount,offset);
        console.log("ItemView");
        var psPropset= new service.EWSMAPI.PropertySet(service.EWSMAPI.BasePropertySet.IdOnly);
        console.log("Propset");
        psPropset.RequestedBodyType = service.EWSMAPI.BodyType.Text;
        psPropset.Add(service.EWSMAPI.ItemSchema.Preview);
        psPropset.Add(service.EWSMAPI.ItemSchema.Subject);
        psPropset.Add(service.EWSMAPI.EmailMessageSchema.Sender);
        psPropset.Add(service.EWSMAPI.ItemSchema.Size);
        psPropset.Add(service.EWSMAPI.ItemSchema.HasAttachments);
        psPropset.Add(service.EWSMAPI.ItemSchema.WebClientReadFormQueryString);
        psPropset.Add(service.EWSMAPI.ItemSchema.DateTimeReceived);
        ItemView.PropertySet = psPropset;
        var FolderItems  = null;
        console.log("Start findItems");
        if(Query == null){
            FolderItems = await service.FindItems(FolderId,ItemView).then(function (response){
                return response;
            });
        }else{
            FolderItems = await service.FindItems(FolderId,Query,ItemView).then(function (response){
                return response;
            });
        }

        return FolderItems;
    }

    module.exports.FindItems = function (service,FolderId,offset,totalCount,Query) {
        return FindItems(service,FolderId,offset,totalCount,Query);
    }

}());

