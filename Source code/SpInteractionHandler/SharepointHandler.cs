using Microsoft.SharePoint.Client;
using MSDN.Samples.ClaimsAuth;
using SP.SpCommonFun;
using SpInteractionHandler.SharepointUpdateHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpInteractionHandler
{
    public class SharepointHandler : IDomainUpdateHandler
    {
        public string TargetSite { get; set; }

        public ContextHandler contextHandler = null;

        public SharepointHandler(string targetURL)
        {
            this.TargetSite = targetURL;

            this.contextHandler = new ContextHandler(targetURL);          
           
        }

        public bool TestConnectivity()
        {
            try
            {
                using (var ctx = this.GetContext())
                {
                    if(ctx == null)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
            
        }

        public List<string> GetAllUpdateableItems()
        {
            using (var ctx = this.GetContext())
            {
                try
                {
                    ListCollection listItems = ctx.Web.Lists;
                    ctx.Load(listItems);
                    ctx.ExecuteQuery();

                    return listItems.Select(o => o.Title).OrderBy(o=> o).ToList();
                }
                catch (Exception ex)
                {
                    AppLogger.Error(ex, "Unable to retrieve All list items from Sharepoint url: " + this.TargetSite);
                    throw;
                }                
            }
        }
        

        public TableListItem GetListItem(string listName)
        {
            try
            {
                using(var ctx = this.GetContext())
                {
                    var listCols = ctx.Web.Lists.GetByTitle(listName);
                    ctx.Load(listCols.Fields);
                    ctx.ExecuteQuery();

                    TableListItem item = new TableListItem();
                    item.ListName = listName;

                    

                    foreach (var field in listCols.Fields.Where(o=> o.FromBaseType == false || 
                    string.Equals(o.EntityPropertyName, "Title", StringComparison.CurrentCultureIgnoreCase)).OrderBy(o=> o.Title))
                    {
                        var columnProp = new TableListColumn()
                        {
                            ColumnName = field.EntityPropertyName,
                            ColumnDisplayName = field.Title,
                            ColumnType = this.GetColumnType(field.FieldTypeKind),
                            ProviderSpecificProperty = field
                        };

                        item.ListColumns.Add(columnProp);
                    }

                    //item.ListColumns = item.ListColumns.OrderBy(o => o.ColumnDisplayName).ToList();
                    return item;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Error retrieving table column names for list: " + listName);
                throw;
            }
        }

        private Type GetColumnType(FieldType fieldType)
        {
            switch (fieldType)
            {
                case FieldType.Text:
                case FieldType.Note:
                    return typeof(string);
                case FieldType.Number:
                    return typeof(double);
                case FieldType.DateTime:
                    return typeof(DateTime);
                case FieldType.Choice:
                    return typeof(string);
                default:
                    return typeof(DateTime);
            }
        }


        private ClientContext GetContext()
        {
            try
            {
                return this.contextHandler.GetContext();
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Exception encountered when retrieving the user context");
                throw;
            }            
        }

        public void UpdateListToSource(string listName, List<IWrappedDataItem> dataToUpdate, TableListColumn[] headerCols, TableListColumn[] primaryKeyCols, IReportProgress reportProgressMethod)
        {
            SPListUpdateHandler listUpdater = new SPListUpdateHandler(this.contextHandler, reportProgressMethod, listName, headerCols, primaryKeyCols);
            listUpdater.PerformAction(dataToUpdate);
        }

        private string GetListURL()
        {
            return this.TargetSite + "_api/web/lists";
        }
    }
    
    public class ContextHandler
    {
        string targetSite = string.Empty;

        public ContextHandler(string targetSite)
        {
            this.targetSite = targetSite;
        }
        public ClientContext GetContext()
        {
            var context =  ClaimClientContext.GetAuthenticatedContext(this.targetSite);
            context.Credentials = System.Net.CredentialCache.DefaultCredentials;
            return context;
        }
    }   
}
