using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContentTypeExtractor.Contracts
{
    interface ISharePointOnlineManager
    {
        
        bool ConnectToSharePoint(string username, string password, out string result);
        List<string> GetContentTypesName(out string result);
        List<string> GetSiteColumnsName(out string result);
        string CreateContentType(string contentTypeName, string parentName = "Item");
        string CreateSiteColumn(string siteColumnName, string contentTypeName, string type, bool isRequired, string group, bool isHidden = false);
        List<string> GetLists(out string result);
        string AddContentTypeToListByNames(string contentTypeName, string listName);
        string DeleteSiteColumnFromContentType(string contentTypeName, string fieldTitle, string fieldType);
        string DeleteContentTypesWithName(string contentType);
        string CreateLibrary(string Title, string Desc);
        //string DeleteList(string contentTypeName, string listName);
    }
}
