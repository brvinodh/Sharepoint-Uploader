using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SpInteractionHandler.SharepointUpdateHandlers.SPListUpdateHandler;

namespace SpInteractionHandler
{
    public interface IReportProgress
    {
        void ReportProgress(int numberOfRecordsRead = -1, int numberOfRecordsUpdated = -1);
    }
    public interface IDomainUpdateHandler
    {        
        List<string> GetAllUpdateableItems();

        bool TestConnectivity();
       
        TableListItem GetListItem(string listName);

        void UpdateListToSource(string listName, List<IWrappedDataItem> dataToUpdate, TableListColumn[] headerCols, TableListColumn[] primaryKeyCols, IReportProgress reportProgressMethod);
    }

    public interface IWrappedDataItem
    {
        dynamic InputData { get; set; }

        SPAction UpdateMode { get; set; }

        UpdateStatus UpdateStatus { get; set; }
    }


    public enum UpdateStatus
    {
        None,

        Success, 

        Failure, 

        Inprogress
    }    

    public class TableListColumn
    {
        public string ColumnName { get; set; }

        public string ColumnDisplayName { get; set; }

        public Type ColumnType { get; set; }

        /// <summary>
        /// Gets or sets the provider specific property. This property would contain the information about the column 
        /// as provided by the Provider, For example if the backend to update is Sharepoint then this would contain 
        /// the Microsoft.SharePoint.Client.Field instance as defined by the Sharepoint framework.
        /// </summary>
        /// <value>
        /// The provider specific property.
        /// </value>
        public object ProviderSpecificProperty { get; set; }

        public bool IsPrimaryKeyCol { get; set; }
    }

    public class TableListItem
    {
        public string ListName { get; set; }

        public List<TableListColumn> ListColumns { get; set; } = new List<TableListColumn>();
    }

}
