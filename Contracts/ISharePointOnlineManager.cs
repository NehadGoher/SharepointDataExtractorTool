using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContentTypeExtractor.Contracts
{
    interface ISharePointOnlineManager
    {
        
        bool ConnectToSharePoint(string username, string password, out string error);
        List<string> GetContentTypesName(out string error);
        List<string> GetSiteColumnsName(out string error);
        string CreateContentType(string contentTypeName, string parentName = "Item");
        string CreateSiteColumn(string siteColumnName, string contentTypeName, string type, bool isRequired, string group, bool isHidden = false);
        List<string> GetLists(out string error);
        string AddContentTypeToListByNames(string contentTypeName, string listName);
        string DeleteSiteColumnFromContentType(string contentTypeName, string fieldTitle, string fieldType);
        string DeleteContentTypesWithName(string contentType);
    }
}
