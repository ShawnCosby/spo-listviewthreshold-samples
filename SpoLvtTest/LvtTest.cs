using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Xml.Linq;

namespace SpoLvtTest
{
    internal class LvtTest
    {
        private readonly string _spoUrl;
        private readonly string _spoUsername;
        private readonly string _spoPassword;

        public LvtTest(string spoUrl, string spoUsername, string spoPassword)
        {
            _spoUrl = spoUrl;
            _spoUsername = spoUsername;
            _spoPassword = spoPassword;
        }

        public void TestList(string listTitle)
        {
            using (var ctx = ClientContextUtility.GetClientContext(_spoUrl, _spoUsername, _spoPassword))
            {
                var web = ctx.Web;
                LoadWeb(web);

                var list = web.Lists.GetByTitle(listTitle);
                var query = new CamlQuery()
                {
                    ViewXml = new XDocument(new XElement("View", new XAttribute("Scope", "RecursiveAll"),
                                            new XElement("RowLimit", 1000))).ToString()
                };

                do
                {
                    LoadList(list);

                    ListItemCollection listItems = list.GetItems(query);
                    ctx.Load(listItems);
                    LoadListItemCollection(listItems);

                    ctx.ExecuteQuery();

                    Console.WriteLine("List '{0}' has a root folder item count of {1}. Will Trigger LVT? {2}", list.Title, list.RootFolder.ItemCount, list.RootFolder.ItemCount > 5000);

                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                } while (query.ListItemCollectionPosition != null);
            }
        }

        public void TestListItem(string listTitle, int listItemID)
        {
            using (var ctx = ClientContextUtility.GetClientContext(_spoUrl, _spoUsername, _spoPassword))
            {
                var web = ctx.Web;
                LoadWeb(web);

                // TO SHAWN: LoadForRecordizeAsync executes the query once before loading list/item
                // Does not throw LVT
                ctx.ExecuteQuery();

                var list = web.Lists.GetByTitle(listTitle);
                // TO SHAWN: LoadForRecordizeAsync loads the list in the context and then loads it again with the specific properties
                //ctx.Load(list);
                LoadList(list);

                // KDD: I added an ExecuteQuery here (not in our original code) to prove that LVT will throw without loading an item
                ctx.ExecuteQuery();

                var listItem = list.GetItemById(listItemID);
                // TO SHAWN: LoadForRecordizeAsync loads the list item in the context and then loads it again with the specific properties
                ctx.Load(listItem);
                LoadListItem(listItem);

                ctx.ExecuteQuery();
            }
        }

        private static void LoadWeb(Web web)
        {
            web.Context.Load(web,
                w => w.Lists
                        .Where(l =>
                            !l.NoCrawl &&
                            !l.Hidden &&
                            !l.IsApplicationList &&
                            l.BaseType == BaseType.DocumentLibrary &&
                            !(l.ListItemEntityTypeFullName == "SP.Data.AppPackagesListItem" || l.ListItemEntityTypeFullName == "SP.Data.FormServerTemplatesItem"))
                        .Include(
                            l => l.RootFolder.Properties,
                            l => l.RootFolder.ServerRelativeUrl,
                            l => l.Id),
                w => w.ServerRelativeUrl,
                w => w.SiteUsers.Include(
                    user => user.Id,
                    user => user.Email),
                w => w.Title,
                w => w.Url,
                w => w.Id);
        }

        private static void LoadList(List list)
        {
            list.Context.Load(list,
                l => l.ContentTypes.Include(
                    contentType => contentType.Id,
                    contentType => contentType.Name,
                    contentType => contentType.Fields
                                                .Where(f => !f.Hidden)
                                                .Include(
                                                    field => field.FieldTypeKind,
                                                    field => field.InternalName,
                                                    field => field.Title)),
                l => l.RootFolder.Name,
                l => l.RootFolder.ItemCount,
                l => l.RootFolder.ServerRelativeUrl,
                l => l.RootFolder.Folders.Include(
                    folder => folder.Name,
                    folder => folder.ServerRelativeUrl),
                l => l.ItemCount,
                l => l.Title);
        }

        private static void LoadListItemCollection(ListItemCollection listItems)
        {
            if (listItems == null)
            {
                throw new ArgumentNullException(nameof(listItems));
            }

            listItems.Context.Load(listItems,
                items => items.Include(
                    item => item.ParentList.Id,
                    item => item.ParentList.RootFolder.ServerRelativeUrl,
                    item => item.ParentList.ParentWeb.Id,
                    item => item.ContentType.Id,
                    item => item.ContentType.Name,
                    item => item.DisplayName,
                    item => item.FileSystemObjectType,
                    item => item.Folder.ServerRelativeUrl,
                    item => item.File.CheckedOutByUser,
                    item => item.File.Exists,
                    item => item.File.Name,
                    item => item.File.ServerRelativeUrl,
                    item => item.File.TimeCreated,
                    item => item.File.TimeLastModified,
                    item => item.File.Title,
                    item => item.File.Length));
        }

        private static void LoadListItem(ListItem listItem)
        {
            if (listItem == null)
            {
                throw new ArgumentNullException(nameof(listItem));
            }

            listItem.Context.Load(listItem,
                item => item.ParentList.Id,
                item => item.ParentList.RootFolder.ServerRelativeUrl,
                item => item.ParentList.ParentWeb.Id,
                item => item.ContentType.Id,
                item => item.ContentType.Name,
                item => item.DisplayName,
                item => item.FileSystemObjectType,
                item => item.Folder.ServerRelativeUrl,
                item => item.File.CheckedOutByUser,
                item => item.File.Exists,
                item => item.File.Name,
                item => item.File.ServerRelativeUrl,
                item => item.File.TimeCreated,
                item => item.File.TimeLastModified,
                item => item.File.Title,
                item => item.File.Length);
        }
    }
}
