using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Xml.Linq;

namespace SpoLvtTest
{
    public static class ListViewThresholdExamples
    {
        public static void TestLvt_SimpleExample(ILogger logger, ClientContext context, string listTitle)
        {
            var list = context.Web.Lists.GetByTitle(listTitle);

            context.Load(list,
                l => l.Title,
                //Including sub-folders will trigger a LVT exception when the root folder has 5,000 or more items
                //SPO work-around: create at least one subfolder and move all items into it so that the root folder holds no more than 5,000 items
                l => l.RootFolder.Folders.Include(f => f.Name));

            context.ExecuteQueryWithIncrementalRetry(throttlingLogMsg => logger.LogWarning(throttlingLogMsg));

            string csomListTitle = list.Title;
            string firstSubFolderName = list.RootFolder.Folders.FirstOrDefault()?.Name;

            System.Diagnostics.Debug.Assert(!string.IsNullOrEmpty(csomListTitle), "Did not retrieve list title property!");
            System.Diagnostics.Debug.Assert(!string.IsNullOrEmpty(firstSubFolderName), "Did not retrieve the first subfolder's Name property!");
        }

        public static void TestLvt_EnumerateListItems(ILogger logger, ClientContext context, string listTitle)
        {
            LoadWeb(context.Web);

            var list = context.Web.Lists.GetByTitle(listTitle);

            LoadList(list);

            var query = new CamlQuery()
            {
                ViewXml = new XDocument(new XElement("View", new XAttribute("Scope", "RecursiveAll"),
                                        new XElement("RowLimit", 1000))).ToString()
            };

            bool listHasMoreItems = true;

            while (listHasMoreItems)
            {
                ListItemCollection listItems = list.GetItems(query);
                context.Load(listItems);
                LoadListItemCollection(listItems);

                context.ExecuteQueryWithIncrementalRetry(throttlingLogMsg => logger.LogWarning(throttlingLogMsg));

                logger.LogInformation("List '{0}' has a root folder item count of {1}. Will Trigger LVT? {2}",
                                      list.Title,
                                      list.RootFolder.ItemCount,
                                      list.RootFolder.ItemCount > 5000);

                query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;

                listHasMoreItems = query.ListItemCollectionPosition != null;
            };

        }

        public static void TestLvt_GetSingleListItem(ILogger logger, ClientContext context, string listTitle, int listItemID)
        {
            LoadWeb(context.Web);

            var list = context.Web.Lists.GetByTitle(listTitle);

            LoadList(list);

            var listItem = list.GetItemById(listItemID);

            LoadListItem(listItem);

            context.ExecuteQueryWithIncrementalRetry(throttlingLogMsg => logger.LogWarning(throttlingLogMsg));
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
                            !(l.ListItemEntityTypeFullName == "SP.Data.AppPackagesListItem" ||
                              l.ListItemEntityTypeFullName == "SP.Data.FormServerTemplatesItem"))
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

            #region Same behavior as the simple example. FIX: subfolder and reorg the root folder so that it holds no more than 5,000 listitems
                        l => l.RootFolder.Folders.Include(
                    folder => folder.Name,
                    folder => folder.ServerRelativeUrl),
            #endregion

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
