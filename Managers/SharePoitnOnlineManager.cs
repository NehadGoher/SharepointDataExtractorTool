using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using System.Security;
using System.Text.RegularExpressions;
using ContentTypeExtractor.Contracts;
using System.Reflection;

namespace ContentTypeExtractor
{
    public sealed class SharePoitnOnlineManager : ISharePointOnlineManager
    {

        List<string> ContentTypes = new List<string>();
        List<string> Fields = new List<string>();
        List<string> spList = new List<string>();

        SP.ClientContext context = null;
        private SP.ContentTypeCollection cts;
        private SP.FieldCollection flds;
        private SP.ListCollection lst;

        public SharePoitnOnlineManager(string url)=> context = new SP.ClientContext(url);

        private void SetCredentials(string username, string password)
        {
            var secureStr = new SecureString();
            foreach (char c in password)
            {
                secureStr.AppendChar(c);
            }
            context.Credentials = new SP.SharePointOnlineCredentials(username, secureStr);
        }

        /// <summary>
        ///  connect to share point online from url pathed in ctor and 
        ///  username and pass pass in the function
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="result"> what happens during process</param>
        /// <returns></returns>
        public bool ConnectToSharePoint(string username, string password,out string result)
        {
            try
            {
                result = String.Empty;
                if (username.Length >0 && password.Length > 0)
                {
                    SetCredentials(username, password);
                    context.Load(context.Web);
                    result = $"Connected to {context.Url} using username : {username}\n";
                }
                else
                {
                    result = "Please set username and password to connect";
                }
                return true;
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
                return false;
            }
        }

        /// <summary>
        /// load from context content types and return list of their names
        /// </summary>
        /// <param name="result"></param>
        /// <returns>list of content type name</returns>
        public List<string> GetContentTypesName(out string result)
        {
            try
            {
                context.Load(context.Web.ContentTypes);
                context.ExecuteQuery();
                cts = context.Web.ContentTypes;
                ContentTypes = cts.ToList().Select(s => s.Name).ToList();
                result = "Content types retrived Successfully \n";
                return ContentTypes;
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
                return ContentTypes;
            }
        }

        /// <summary>
        /// load site column from context 
        /// </summary>
        /// <param name="result"> out the result of what happen</param>
        /// <returns>list of string from site columns</returns>
        public List<string> GetSiteColumnsName(out string result)
        {
            try
            {
                context.Load(context.Web.Fields);
                context.ExecuteQuery();
                flds = context.Web.Fields;
                Fields = flds.ToList().Select(s => s.InternalName).ToList();
                result = "Site column retrived Successfully \n";
                return Fields;
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
                return Fields;
            }
        }

        /// <summary>
        ///  create content type by name and name of parent
        /// </summary>
        /// <param name="contentTypeName"></param>
        /// <param name="parentName"></param>
        /// <returns> of result or result of what happens </returns>
        public string CreateContentType(string contentTypeName, string parentName = "Item")
        {
            try
            {
                string name = Regex.Replace(contentTypeName, @"\s+", "");
                if (ContentTypes.Contains(name))
                {
                    return "ContentType : " + contentTypeName + " already Exists \n";
                }
                else if (ContentTypes.Count > 0)
                {

                    SP.ContentType parent = null;
                   
                    parent = cts.ToList().Where(s => s.Name == parentName).FirstOrDefault();
                    if (parent != null)
                    {
                        cts.Add(new SP.ContentTypeCreationInformation {  Name = Regex.Replace(name, @"\s+", ""), ParentContentType = parent});
                    }
                    context.ExecuteQuery();
                    ContentTypes.Add(Regex.Replace(name, @"\s+", ""));
                    return $"ContentType : {name}, parent : {parentName}\n";
                }
                else
                {
                    return "No Available Content Types";
                }
            }
            catch(Exception ex)
            {
                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }
        }

        /// <summary>
        /// Create SiteColumn with name of site and name of content type and type
        /// </summary>
        /// <param name="siteColumnName"></param>
        /// <param name="contentTypeName"></param>
        /// <param name="type"></param>
        /// <param name="isRequired"></param>
        /// <param name="group"></param>
        /// <param name="isHidden"></param>
        /// <returns>of result or result of what happens</returns>
        public string CreateSiteColumn(string siteColumnName, string contentTypeName, string type, bool isRequired, string group, bool isHidden = false)
        {
            try
            {
                string name = Regex.Replace(siteColumnName, @"\s+", "");
                if (Fields.Contains(name))
                {
                    return "Site Column : " + name + " already Exists\n";
                }
                else
                {
                    /// check if content type for site column is exists
                    if (ContentTypes.Contains(Regex.Replace(contentTypeName, @"\s+", "")))
                    {

                        string fieldAsXml =
                            $@"<Field ID=""{{{Guid.NewGuid()}}}"" Name=""{Regex.Replace(name, @"\s+", "")}"" " +
                            $@"DisplayName=""{name}"" " +
                            $@"Type=""{type}"" Hidden=""{false}"" " +
                            $@"Required=""{isRequired}"" Group=""{group}"" />";

                        SP.Field _field = flds.AddFieldAsXml(fieldAsXml, true, SP.AddFieldOptions.DefaultValue);
                        context.Load(_field);
                        SP.FieldLinkCreationInformation link = new SP.FieldLinkCreationInformation { Field = _field };
                        /// link  column to content type
                        var cntTyp = cts.Where(s => s.Name ==contentTypeName).FirstOrDefault();
                        if (cntTyp != null)
                        {
                            var newRefField = cntTyp.FieldLinks.Add(link);
                            newRefField.Hidden = false;
                            newRefField.Required = isRequired;
                        }
                        cntTyp.Update(true);
                        context.ExecuteQuery();
                        Fields.Add(Regex.Replace(name, @"\s+", ""));
                        return $"Site Column : {name} , for content type : {contentTypeName}\n";
                    }
                    else
                    {
                        return $"Content type : {contentTypeName} not exists \n";

                    }
                }
            }
            catch(Exception ex)
            {
                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }
           
        }
        
        /// <summary>
        /// get all lists in sharepoint
        /// </summary>
        /// <param name="result"></param>
        /// <returns>list of list</returns>
        public List<string> GetLists(out string result)
        {
            try
            {
                result = String.Empty;
                context.Load(context.Web.Lists);
                context.ExecuteQuery();
                lst = context.Web.Lists;
                spList = lst.ToList().Select(s => s.Title).ToList();
                return spList;
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
                return spList;
            }
        }
        
        /// <summary>
        /// add content type to a list in share point
        /// </summary>
        /// <param name="contentTypeName"></param>
        /// <param name="listName"></param>
        /// <returns>result happens in adding</returns>
        public string AddContentTypeToListByNames( string contentTypeName, string listName)
        {
            try
            {
                if(!string.IsNullOrWhiteSpace(contentTypeName) && !string.IsNullOrWhiteSpace(listName))
                {
                    var list = lst.ToList().Where(s => s.Title == Regex.Replace(listName, @"\s+", "")).FirstOrDefault();
                    if (list == null) return $"Library : {listName} can't be found\n";
                    var contentType = cts.ToList().Where(s => s.Name == Regex.Replace(contentTypeName, @"\s+", "")).FirstOrDefault();
                    if (contentType == null) return $"Content type : {contentTypeName} can't be found\n";
                    list.ContentTypesEnabled = true;
                    list.Update();
                    list.ContentTypes.AddExistingContentType(contentType);
                    context.ExecuteQuery();
                    return $"Library : {listName} Linked to {contentTypeName}\n";
                }
                else
                {
                    return "ContentTypeName and list name must have value \n";
                }
            }catch(Exception ex)
            {

                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }
        }

        /// <summary>
        /// Delete site column by content type and site name and type
        /// </summary>
        /// <param name="contentTypeName"></param>
        /// <param name="fieldTitle"></param>
        /// <param name="fieldType"></param>
        /// <returns>of what happen during the process fail or sucess</returns>
        public string DeleteSiteColumnFromContentType(string contentTypeName, string fieldTitle, string fieldType)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(contentTypeName) && !string.IsNullOrWhiteSpace(fieldTitle) && !string.IsNullOrWhiteSpace(fieldType))
                {
                    SP.Field spField = flds.FirstOrDefault(x => (x.EntityPropertyName == fieldTitle && x.TypeAsString.Contains(fieldType)));
                    
                    SP.ContentType ct = cts.FirstOrDefault(x => x.Name == contentTypeName);
                    if (spField != null)
                    {
                        //context.Load(spField);
                        //context.ExecuteQuery();
                        if (ct != null) { 
                            SP.FieldLinkCollection links = ct.FieldLinks;
                            SP.FieldLink fieldLinkToRemove = links.GetById(spField.Id);
                            if (fieldLinkToRemove != null)
                            {
                                // fieldLinkToRemove.RefreshLoad();
                                // context.ExecuteQuery();
                                fieldLinkToRemove.DeleteObject();
                                //context.Load(links);
                            }
                            //context.ExecuteQuery();
                            ct.Update(true); //push changes
                        }
                        //context.Load(spField);
                        if (spField.CanBeDeleted)
                        {
                            spField.DeleteObject();
                            
                        }
                        else
                        {
                            return $"{fieldTitle} can't be deleted \n";
                        }
                        //context.Load(ct);
                        context.ExecuteQuery();
                        Fields.Remove(fieldTitle);
                        return $"{fieldTitle} deleted from {contentTypeName} \n";
                    }
                    else
                    {
                        return $"{fieldTitle} not found \n";
                    }
                    
                }
                return $"content type name and sitecolumn must have value \n";
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }
        }

        /// <summary>
        ///  Delete content type by name
        /// </summary>
        /// <param name="contentTypeName"></param>
        /// <returns>of what happen during the process fail or sucess</returns>
        public string DeleteContentTypesWithName(string contentTypeName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(contentTypeName))
                {
                    SP.ContentType contentType = cts.Where(s => s.Name == contentTypeName).FirstOrDefault();
                    //// can i remove item ??
                    if (contentType != null)
                    {
                        contentType.DeleteObject();
                        context.ExecuteQuery();
                        ContentTypes.Remove(contentTypeName);
                        return $"{contentTypeName} deleted successfully \n";
                    }
                    else
                    {
                        return $"{contentTypeName} doen't exists \n";
                    }

                }
                return "content Type Name must have value \n";
            }
            catch (Exception ex)
            {
                // return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n"; 
                return $"{contentTypeName} failed to be deleted \n"; 
            }

        }

        /// <summary>
        /// Create library with name
        /// </summary>
        /// <param name="Title"></param>
        /// <param name="Desc"></param>
        /// <returns></returns>
        public string CreateLibrary(string Title, string Desc = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Title))
                {
                    if (spList.Any(l=> l==Title))
                    {
                        return $"Library : {Title} already exists\n";
                    }
                    else
                    {
                        SP.ListCreationInformation newLibrary = new SP.ListCreationInformation()
                        {
                            Title = Title,
                            Description = Desc,
                            TemplateType = (int)SP.ListTemplateType.DocumentLibrary
                        };  
                        lst.Add(newLibrary);
                        context.Load(lst);
                        context.ExecuteQuery();
                        spList.Add(Title);
                        return $"Library : {Title} Created successfully\n";
                    }
                }
                return "Must Enter Valid title list and desc\n";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }

        }

        public string DeleteList( string listName)
        {
            try
            {
                if ( !string.IsNullOrWhiteSpace(listName))
                {
                     var lib =lst.ToList().Where(l=>l.Title == listName).FirstOrDefault();
                    if (lib == null) return $"Library : {listName} doesn't exist\n";
                    lib.DeleteObject();
                    context.ExecuteQuery();
                    return $"Library : {listName} deleted successfully\n";
                }
                else
                {
                    return "Library name must have value \n";
                }
            }
            catch (Exception ex)
            {

                return ex.Message.ToString() + $" in {MethodBase.GetCurrentMethod()}\n";
            }
        }
    }
}
