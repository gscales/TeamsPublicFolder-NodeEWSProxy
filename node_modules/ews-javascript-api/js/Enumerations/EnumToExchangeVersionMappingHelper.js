"use strict";
/**custom created to simplify creation of above Enum(s) to ExchangeVersion mapping in EwsUtil, There is no c# like Attribute typesystem and reflection available */
var EnumToExchangeVersionMappingHelper;
(function (EnumToExchangeVersionMappingHelper) {
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["WellKnownFolderName"] = 0] = "WellKnownFolderName";
    /**Item Traversal */
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["ItemTraversal"] = 1] = "ItemTraversal";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["ConversationQueryTraversal"] = 2] = "ConversationQueryTraversal";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["FileAsMapping"] = 3] = "FileAsMapping";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["EventType"] = 4] = "EventType";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["MeetingRequestsDeliveryScope"] = 5] = "MeetingRequestsDeliveryScope";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["ViewFilter"] = 6] = "ViewFilter";
    EnumToExchangeVersionMappingHelper[EnumToExchangeVersionMappingHelper["MailboxType"] = 7] = "MailboxType";
})(EnumToExchangeVersionMappingHelper = exports.EnumToExchangeVersionMappingHelper || (exports.EnumToExchangeVersionMappingHelper = {}));
