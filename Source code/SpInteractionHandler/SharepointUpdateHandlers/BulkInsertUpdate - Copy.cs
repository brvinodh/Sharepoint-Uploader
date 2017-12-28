using Microsoft.SharePoint.Client;
using SP.SpCommonFun;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpInteractionHandler.SharepointUpdateHandlers
{
    
    public class SPListUpdateHandler : IUpdateMode
    {
        public int UpdateModeID { get; set; }

        public string UpdateModeText { get; set; }

        public string ModeDescription { get; set; }

        private ContextHandler contextHandler { get; set; }

        public StreamWriter progressLogStream;

        public List<TableListColumn> ListFields { get; set; } = new List<TableListColumn>();

        public List<dynamic> SharepointListData { get; set; } = new List<dynamic>();

        /// <summary>
        /// Gets or sets the batch commit. Indicates the number of records after which a sharepoint commit should be triggered
        /// </summary>
        /// <value>
        /// The batch commit.
        /// </value>
        public int BatchCommit { get; set; } = 100;

        public int NumberOfUpdateThreads { get; set; } = 1;

        /// <summary>
        /// Gets or sets the sp list data to update.
        /// This object would be used in multithreaded environment. We would have many threads which may want to update the 
        /// list to sharepoint; this data structure acts as a input from which data can be read in individual threads.
        /// </summary>
        /// <value>
        /// The sp list data to update.
        /// </value>
        private ConcurrentQueue<dynamic> spListDataToUpdate { get; set; } = new ConcurrentQueue<dynamic>();


        public bool RequestAbort { get; set; } = false;

        private volatile bool allRecordsComparisonComplete = false; 

        public SPListUpdateHandler(ContextHandler contextHandler, StreamWriter logStream)
        {
            this.contextHandler = contextHandler;
        }

        public virtual void PerformAction(string listName, TableListColumn[] headerColumnsFieldsList, TableListColumn[] primaryKeyCols, List<dynamic> dataToUpdate)
        {
            string stepName = "Init";
            try
            {
                this.allRecordsComparisonComplete = false;
                stepName = "Read all items from SP";

                string[] headerCols = headerColumnsFieldsList.Select(o => o.ColumnName).ToArray();
                string[] primaryCols = primaryKeyCols.Select(o => o.ColumnName).ToArray();





                // step 1: Read all records from Sharepoint
                //this.ReadAllItems(listName, headerCols);

                // Step 1a: Initiate threads which run in continous loop; this thread would be a consumer
                // consuming every item prducted by the CompareSourceWithData method
                for (int i = 0; i < this.NumberOfUpdateThreads; i++)
                {
                    Task.Run(() =>
                    {
                        this.ExecuteContinousUpdate(listName, headerCols);
                    });
                }

                // approach B:
                // for every item in the input data query and compare; then add to update queue. 
                this.CompareUpdateInputDataWithServerData(listName, dataToUpdate, primaryCols, headerCols);

                // step 2: Compare Read Records against data
                //this.CompareSourceWithData(this.SharepointListData, dataToUpdate, headerCols, primaryCols);

                // step 3: Insert/Update phase
                //this.UpdateItem(listName, headerCols);
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"Exception when performing update action at {stepName}");
                throw;
            }           
        }

        private void CompareUpdateInputDataWithServerData(string listName, List<dynamic> dataToUpdate, string[] primaryKeyCols, string[] headerCols)
        {
            // read all the records and enquue them to an update queue
            this.ReadMultipleItemsGroupBatch(listName, primaryKeyCols, headerCols, dataToUpdate);

            // signal that all the records have been evaluated.
            this.allRecordsComparisonComplete = true;
        }

        private dynamic ReadSingleItem(string listName, string[] primaryKeyCols, string[] headerCols, IDictionary<string, object> inputdata)
        {
            try
            {
                const int rowLimit = 5000;

                CamlQuery query = new CamlQuery();
                string retrievableFields = string.Empty;
                string whereCondition = string.Empty;

                Array.ForEach(primaryKeyCols, o => whereCondition += 
                $" <Eq><FieldRef Name=\"{o}\" /><Value Type='Text'>{inputdata[o]}</Value></Eq>"                
                );

                Array.ForEach<string>(headerCols, o => retrievableFields += "<FieldRef Name='" + o + "'/>");                

                // set the query to get only 5000 items at one shot
                query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" + retrievableFields + "</ViewFields><RowLimit>" + 5 + "</RowLimit><Query><Where> " + whereCondition+ " </Where></Query></View>";

                using (var ctx = this.contextHandler.GetContext())
                {
                    int recordsRetrieved = 0;
                    int batch = 0;
                    ListItemCollectionPosition position = null;
                    this.ReportProgress($"Reading specific items from Sharepoints for list {listName}; QueryXML: {query.ViewXml}");

                    var oList = ctx.Web.Lists.GetByTitle(listName);

                    do
                    {
                        this.ReportProgress($"Reading batch: {batch}");
                        DateTime dt = DateTime.Now;
                        
                        ListItemCollection listItems = oList.GetItems(query);

                        ctx.Load(listItems);
                        ctx.ExecuteQuery();

                        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: Successful Read of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");

                        // add item to ListItemsOnServer once retrieved so that it can be used for processing
                        position = listItems.ListItemCollectionPosition;
                        query.ListItemCollectionPosition = position;
                        recordsRetrieved = listItems.Count;

                        for (int i = 0; i < recordsRetrieved; i++)
                        {
                            //this.SharepointListData.Add(this.MapFromListItem(listItems[i], headerCols));
                            return this.MapFromListItem(listItems[i], headerCols);
                        }

                    } while (position != null);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Error encountered when querying for input data from sharepoint");
                throw;
            }

            return null;
        }

        private void ReadMultipleItems(string listName, string[] primaryKeyCols, string[] headerCols, List<dynamic> inputdata)
        {
            try
            {                
                int primaryKeysCount = primaryKeyCols.Length;

                CamlQuery query = new CamlQuery();
                string retrievableFields = string.Empty;               
                Array.ForEach<string>(headerCols, o => retrievableFields += "<FieldRef Name='" + o + "'/>");                                              

                using (var ctx = this.contextHandler.GetContext())
                {
                    int inputdataCount = inputdata.Count();
                    var oList = ctx.Web.Lists.GetByTitle(listName);

                    int recordsRetrieved = 0;
                    int batch = 0;
                   
                    for (int iInputIndex = 0; iInputIndex < inputdataCount; iInputIndex++)
                    {
                        string whereCondition = string.Empty;
                        IDictionary<string, object> inputItem = inputdata[iInputIndex];

                        for (int i = 0; i < primaryKeysCount; i++)
                        {
                            string fieldName = primaryKeyCols[i];
                            whereCondition += $"<Eq><FieldRef Name=\"{fieldName}\" /><Value Type='Text'>{inputItem[fieldName]}</Value></Eq>";
                        }

                        // set the query to get only 5000 items at one shot
                        query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" + retrievableFields + "</ViewFields><RowLimit>" + 5 + "</RowLimit><Query><Where> " + whereCondition + " </Where></Query></View>";

                        //this.ReportProgress($"Reading specific items from Sharepoints for list {listName}; QueryXML: {query.ViewXml}");
                                                
                        DateTime dt = DateTime.Now;

                        ListItemCollection listItems = oList.GetItems(query);

                        ctx.Load(listItems);
                        ctx.ExecuteQuery();
                        //Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: Successful Read of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");
                        recordsRetrieved = listItems.Count;

                        for (int i = 0; i < recordsRetrieved; i++)
                        {
                            //this.SharepointListData.Add(this.MapFromListItem(listItems[i], headerCols));
                            this.CompareSingleItem(this.MapFromListItem(listItems[i], headerCols), inputItem, headerCols);
                            // return this.MapFromListItem(listItems[i], headerCols);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Error encountered when querying for input data from sharepoint");
                throw;
            }            
        }

        private void ReadMultipleItemsGroupBatch(string listName, string[] primaryKeyCols, string[] headerCols, List<dynamic> inputdata)
        {
            try
            {
                int primaryKeysCount = primaryKeyCols.Length;

                CamlQuery query = new CamlQuery();
                string retrievableFields = string.Empty;
                Array.ForEach<string>(headerCols, o => retrievableFields += "<FieldRef Name='" + o + "'/>");

                using (var ctx = this.contextHandler.GetContext())
                {
                    int inputdataCount = inputdata.Count();
                    var oList = ctx.Web.Lists.GetByTitle(listName);

                    int recordsRetrieved = 0;
                    int batch = 6;
                    
                    

                    for (int iInputIndex = 0; iInputIndex < inputdataCount; iInputIndex= iInputIndex + batch)
                    {
                        string whereCondition = string.Empty;
                        string[] whereConditions = new string[batch];

                        List<dynamic> inputItems = new List<dynamic>();
                        int nextMaxBatchIndex = iInputIndex + batch;
                        nextMaxBatchIndex = nextMaxBatchIndex < inputdataCount ? nextMaxBatchIndex : inputdataCount;
                        int ctr = 0;
                        for (int j = iInputIndex; j < nextMaxBatchIndex; j++)
                        {
                            IDictionary<string, object> inputItem = inputdata[j];
                            for (int i = 0; i < primaryKeysCount; i++)
                            {
                                string fieldName = primaryKeyCols[i];
                                whereConditions[ctr] += $"<Eq><FieldRef Name=\"{fieldName}\" /><Value Type='Text'>{inputItem[fieldName]}</Value></Eq>";
                            }

                            inputItems.Add(inputItem);
                            ctr++;                            
                        }

                        if (batch > 0)
                        {
                            // the Or Logical operator can only have 2 conditions at one short hence we need 
                            // collapse them to suit 2 conditions
                            List<string> orConditionsList = new List<string>();
                            orConditionsList.AddRange(whereConditions.Where(o=> !string.IsNullOrEmpty(o)));

                            while (true)
                            {
                                List<string> newConditionsList = new List<string>();
                                int condnCount = orConditionsList.Count();
                                for (int orIndex = 0; orIndex < condnCount; orIndex = orIndex +2)
                                {
                                    string temp = string.Empty;
                                    if (orIndex + 1 < condnCount)
                                    {
                                        temp = "<Or>" + orConditionsList[orIndex] + orConditionsList[orIndex + 1] + "</Or>";
                                    }
                                    else
                                    {
                                        temp = orConditionsList[orIndex];
                                    }

                                    newConditionsList.Add(temp);
                                }

                                orConditionsList = newConditionsList;

                                // if there are no more conditions to club then break from the loop
                                if (newConditionsList.Count()== 1)
                                {
                                    whereCondition = newConditionsList.FirstOrDefault();
                                    break;
                                }
                            }
                        }


                        // set the query to get only 5000 items at one shot
                        query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" + retrievableFields + "</ViewFields><RowLimit>" + batch + "</RowLimit><Query><Where> " + whereCondition + " </Where></Query></View>";

                        //this.ReportProgress($"Reading specific items from Sharepoints for list {listName}; QueryXML: {query.ViewXml}");

                        DateTime dt = DateTime.Now;

                        ListItemCollection listItems = oList.GetItems(query);

                        ctx.Load(listItems);
                        ctx.ExecuteQuery();
                        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: Successful Read of batch: {batch} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");
                        recordsRetrieved = listItems.Count;

                        List<dynamic> serverItems = new List<dynamic>();
                        for (int i = 0; i < recordsRetrieved; i++)
                        {
                            //this.SharepointListData.Add(this.MapFromListItem(listItems[i], headerCols));
                            //this.CompareSingleItem(this.MapFromListItem(listItems[i], headerCols), inputItem, headerCols);
                            // return this.MapFromListItem(listItems[i], headerCols);
                            serverItems.Add(this.MapFromListItem(listItems[i], headerCols));
                        }
                        this.CompareSourceWithData(serverItems, inputItems, headerCols, primaryKeyCols);

                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Error encountered when querying for input data from sharepoint");
                throw;
            }
        }

        private void CompareSingleItem(IDictionary<string, object> serverItem, IDictionary<string, object> clientItem, string[] headerCols)
        {
            SPAction actionType = serverItem == null ? SPAction.Insert : SPAction.NoAction;
            int totalFields = headerCols.Count();

            if (serverItem != null)
            {
                bool needUpdate = false;
                for (int colIndex = 0; colIndex < totalFields; colIndex++)
                {
                    string propertyToCompare = headerCols[colIndex];
                    needUpdate = !this.CompareValueOrRefType(serverItem, clientItem, propertyToCompare);
                    if (needUpdate) break;
                }

                if (needUpdate)
                {
                    clientItem["ID"] = serverItem["ID"];
                    actionType = SPAction.Update;
                }
                else
                    actionType = SPAction.NoAction;
            }

            if (actionType != SPAction.NoAction)
            {
                // we enqueue items which needs to be processed into the list action queue
                // On a different thread these queues will be read and processed.
                clientItem["UpdateMode"] = actionType;
                this.spListDataToUpdate.Enqueue(clientItem);
            }
        }

        public void ReadAllItems(string listName, string[] headerCols)
        {
            try
            {
                CamlQuery query = new CamlQuery();
                string retrievableFields = string.Empty;
                Array.ForEach<string>(headerCols, o => retrievableFields += "<FieldRef Name='" + o + "'/>");

                const int rowLimit = 5000;

                // set the query to get only 5000 items at one shot
                query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" + retrievableFields + "</ViewFields><RowLimit>" + rowLimit + "</RowLimit></View>";

                // Sharepoint does not allow to retrieve all records at one shot hence we need to read in batches of 5000
                using (var ctx = this.contextHandler.GetContext())
                {
                    int recordsRetrieved = 0;
                    int batch = 0;
                    ListItemCollectionPosition position = null;
                    this.ReportProgress($"Reading all items from Sharepoints for list {listName}; QueryXML: {query.ViewXml}");

                    do
                    {
                        this.ReportProgress($"Reading batch: {batch}");
                        DateTime dt = DateTime.Now;
                        var oList = ctx.Web.Lists.GetByTitle(listName);
                        ListItemCollection listItems = oList.GetItems(query);
                        
                        ctx.Load(listItems);
                        ctx.ExecuteQuery();

                        Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: Successful Read of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");

                        // add item to ListItemsOnServer once retrieved so that it can be used for processing
                        position = listItems.ListItemCollectionPosition;
                        query.ListItemCollectionPosition = position;
                        recordsRetrieved = listItems.Count;

                        for (int i = 0; i < recordsRetrieved; i++)
                        {
                            this.SharepointListData.Add(this.MapFromListItem(listItems[i], headerCols));
                        }

                    } while (position != null);
                }

                Console.WriteLine("All records read from Sharepoint");
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "There was an error when reading sharepoint records");
                throw;
            }            
        }

        private dynamic MapFromListItem(ListItem listItem, string[] headerCols)
        {
            dynamic tempObj = new System.Dynamic.ExpandoObject();

            var dictionary = (IDictionary<string, object>)tempObj;

            for (int i = 0; i < headerCols.Length; i++)
            {
                string fieldName = headerCols[i];
                dictionary[fieldName] = listItem[fieldName];
            }
            
            // read the id field name as this is mandatory one.
            dictionary["ID"] = listItem["ID"];

            return tempObj;
        }

        private void ReportProgress(string comment)
        {
            // this.progressLogStream.WriteLine(comment);
            Console.WriteLine(comment);
            AppLogger.Info(comment);
        }

        private void CompareSourceWithData(List<dynamic> source, List<dynamic> updateData, string[] headerCols, string[] primaryCols)
        {
            int totalUpdateCount = updateData.Count;

            for (int i = 0; i < totalUpdateCount; i++)
            {
                var item = updateData[i];
                SPAction actionType = this.Compare(source, item, headerCols, primaryCols);

                item.UpdateMode = actionType;

                if(actionType != SPAction.NoAction)
                {
                    // we enqueue items which needs to be processed into the list action queue
                    // On a different thread these queues will be read and processed.
                    this.spListDataToUpdate.Enqueue(item);
                }
            }

            //this.allRecordsComparisonComplete = true;
        }

        private SPAction Compare(List<dynamic> source, dynamic target, string[] headerCols, string[] primaryCols)
        {
            int totalCount = headerCols.Count();
            int primaryColCount = primaryCols.Length;

            var targetDictObj = (IDictionary<string, object>)target;

            var matchObj = source.Where(o =>
              {               

                  bool matchFound = primaryColCount > 0;

                  for (int pColindex = 0; pColindex < primaryColCount && matchFound; pColindex++)
                  {
                      string propertyName = primaryCols[pColindex];
                      matchFound = this.CompareValueOrRefType(o, targetDictObj, propertyName);
                      
                      if (!matchFound) break;
                  }

                  return matchFound;
              }).FirstOrDefault();


            if (matchObj == null)
            {
                // insert Case
                return SPAction.Insert;
            }
            else
            {
                // check if all the other properties have changed, if yes then need an update
                bool needUpdate = false;
                for (int colIndex = 0; colIndex < totalCount; colIndex++)
                {                    
                    string propertyToCompare = headerCols[colIndex];
                    needUpdate = !this.CompareValueOrRefType(matchObj, targetDictObj, propertyToCompare);                    
                    if (needUpdate) break;
                }

                if (needUpdate)
                {
                    target.ID = matchObj.ID;
                    return SPAction.Update;
                }
                else
                    return SPAction.NoAction;
            }
        }
                
        private void UpdateListItem(ListItem oListItem, IDictionary<string, object> nextItem, string[] fieldNames)
        {
            int totalFields = fieldNames.Length;
            for (int i = 0; i < totalFields; i++)
            {
                string field = fieldNames[i];
                oListItem[field] = nextItem[field];
            }

            oListItem.Update();
        }

        private bool CompareValueOrRefType(IDictionary<string, object> obj1, IDictionary<string, object> obj2, string propertyName)
        {
            object propA = obj1[propertyName];
            object propB = obj2[propertyName];

            bool isValueType = (propA != null) && (propA.GetType() ==  typeof(string) || propA.GetType().IsValueType);

            if (isValueType)
            {
                propA = (propA is DateTime) ? ((DateTime)propA).Date : propA;
                propB = (propB is DateTime) ? ((DateTime)propB).Date : propB;

                bool isSame = propA.Equals(propB);
                return isSame;
            }
            else
            {
                return propA == propB;
            }
        }

        private void ExecuteContinousUpdate(string listName, string[] headerCols)
        {
            try
            {
                using (ClientContext ctx = this.contextHandler.GetContext())
                {
                    if (ctx != null)
                    {
                        var oList = ctx.Web.Lists.GetByTitle(listName);
                        var listCreationinformation = new ListItemCreationInformation();
                        int count = 0;
                        int batch = 0;

                        while (!this.allRecordsComparisonComplete || this.spListDataToUpdate.Count() != 0)
                        {
                            if (this.spListDataToUpdate.Count() == 0)
                                Thread.Sleep(TimeSpan.FromSeconds(5));

                            dynamic nextItem = null;
                            if (this.spListDataToUpdate.TryDequeue(out nextItem))
                            {
                                ListItem oListItem = null;
                                if (nextItem.UpdateMode == SPAction.Insert)
                                {
                                    oListItem = oList.AddItem(listCreationinformation);
                                }
                                else if (nextItem.UpdateMode == SPAction.Update)
                                {
                                    oListItem = oList.GetItemById(nextItem.ID);
                                }

                                 this.UpdateListItem(oListItem, nextItem, headerCols);

                                count++;
                                
                                // commit for every x records ex: 100 records
                                if (count >= this.BatchCommit)
                                {
                                    DateTime dt = DateTime.Now;
                                    try
                                    {
                                        ctx.ExecuteQuery();
                                        nextItem.UpdateMode = "ChangedAndUpdated";
                                    }
                                    catch (Exception ex)
                                    {
                                        AppLogger.Error(ex, "Error updating a specific item");
                                        throw;
                                    }
                                    
                                    //Console.WriteLine($"In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}: at count: {this.TotalCount++}");

                                    Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}, Successful update of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");
                                    this.ReportProgress($"{DateTime.Now.ToString("HH:mm:ss")}: In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}, Successful update of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");
                                    count = 0;
                                }
                            }
                        }

                        // if there are some pending items to be updated then update them
                        if (count != 0)
                        {
                            ctx.ExecuteQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"Unexpected error in thread: {Thread.CurrentThread.ManagedThreadId}");
                this.ReportProgress("Error encountered when updating the records to sharepoint");                
            }              
        }

        public enum SPAction
        {
            Insert,

            Delete,

            Update,

            NoAction
        }
    }
}
