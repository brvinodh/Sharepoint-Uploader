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
    
    public class SPListUpdateHandler
    {
        public int UpdateModeID { get; set; }

        public string UpdateModeText { get; set; }

        public string ModeDescription { get; set; }

        private ContextHandler contextHandler { get; set; }

        public IReportProgress progressReporter;
        
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
        private ConcurrentQueue<IWrappedDataItem> spListDataToUpdate { get; set; } = new ConcurrentQueue<IWrappedDataItem>();
        
        public bool RequestAbort { get; set; } = false;

        private volatile bool allRecordsComparisonComplete = false;

        #region ListRelatedProperties
        public string ListName { get; set; }

        public TableListColumn[] ListColumnsToUpdate { get; set; }

        public TableListColumn[] ListPrimaryKeyColumns { get; set; }

        private int listColumnsCount = 0;
        #endregion

        public SPListUpdateHandler(ContextHandler contextHandler, IReportProgress reportProgress, string listName, TableListColumn[] listHeaderColumns, TableListColumn[] liskPrimaryKeyColumns)
        {
            this.contextHandler = contextHandler;
            this.progressReporter = reportProgress;
            this.ListName = listName;
            this.ListColumnsToUpdate = listHeaderColumns;
            this.ListPrimaryKeyColumns = liskPrimaryKeyColumns;
            this.listColumnsCount = this.ListColumnsToUpdate.Count();
        }

        public virtual void PerformAction(List<IWrappedDataItem> dataToUpdate)
        {
            string stepName = "Init";
            try
            {
                this.allRecordsComparisonComplete = false;
                stepName = "Read all items from SP";

                // Step 1a: Initiate threads which run in continous loop; this thread would be a consumer
                // consuming every item prducted by the CompareSourceWithData method
                List<Task> taskLIst = new List<Task>();
                for (int i = 0; i < this.NumberOfUpdateThreads; i++)
                {
                    taskLIst.Add(Task.Run(() =>
                   {
                       AppLogger.Info($"Started a Thread: {Thread.CurrentThread.ManagedThreadId} to update list: {this.ListName}");
                       this.ExecuteContinousUpdate();
                   }));
                }

                // Step 2:
                // for every item in the input data query and compare; then add to update queue. 
                this.CompareUpdateInputDataWithServerData(dataToUpdate);

                Task.WaitAll(taskLIst.ToArray());

            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"Exception when performing update action at {stepName}");
                throw;
            }           
        }

        private void CompareUpdateInputDataWithServerData(List<IWrappedDataItem> dataToUpdate)
        {
            // read all the records and enquue them to an update queue
            this.ReadMultipleItemsGroupBatch(dataToUpdate);

            // signal that all the records have been evaluated.
            this.allRecordsComparisonComplete = true;
        }
        private void ReadMultipleItemsGroupBatch(List<IWrappedDataItem> inputdata)
        {
            try
            {                
                CamlQuery query = new CamlQuery();
                string retrievableFields = string.Empty;

                // create a string with list of fieldnames to be retrieved
                Array.ForEach(this.ListColumnsToUpdate, o => retrievableFields += "<FieldRef Name='" + o.ColumnName + "'/>");

                using (var ctx = this.contextHandler.GetContext())
                {
                    int inputdataCount = inputdata.Count();
                    var oList = ctx.Web.Lists.GetByTitle(this.ListName);

                    int recordsRetrieved = 0;
                    int batch = 30;  

                    for (int iInputIndex = 0; iInputIndex < inputdataCount; iInputIndex++)
                    {
                        string whereCondition = string.Empty;

                        // the records are queried in batches instead of single query
                        // this would increase performance
                        // get the next xx number of records from the input list where xx is the value specified by batch variable
                        List<IWrappedDataItem> itemsForQuery = new List<IWrappedDataItem>();
                        int tempI = 0;
                        
                        while (tempI < batch && iInputIndex < inputdataCount)
                        {
                            itemsForQuery.Add(inputdata[iInputIndex]);
                            tempI++;
                            iInputIndex++;
                        }

                        // post to the last item which was not added
                        iInputIndex--;                        

                        whereCondition = this.GetWhereQueryConditionFor(itemsForQuery);

                        // set the query to get only 5000 items at one shot
                        query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" + retrievableFields + "</ViewFields><RowLimit>" + batch + "</RowLimit><Query><Where> " + whereCondition + " </Where></Query></View>";

                        //this.ReportProgress($"Reading specific items from Sharepoints for list {listName}; QueryXML: {query.ViewXml}");

                        DateTime dt = DateTime.Now;

                        ListItemCollection listItems = oList.GetItems(query);

                        ctx.Load(listItems);
                        ctx.ExecuteQuery();
                        string log = $"{DateTime.Now.ToString("HH:mm:ss")}: Successful Read of batch: {batch} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds";
                        // Console.WriteLine();

                        this.ReportProgress(log, recordsRead: iInputIndex + 1);

                        recordsRetrieved = listItems.Count;
                        
                        this.CompareSourceWithData(listItems.ToList(), itemsForQuery);

                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Error encountered when querying for input data from sharepoint");
                throw;
            }
        }

        private string GetWhereQueryConditionFor(List<IWrappedDataItem> userInputData)
        {
            int count = userInputData.Count();
            int primaryKeysCount = this.ListPrimaryKeyColumns.Length;

            // we will have condition for each input item
            // each input item condition will be clubbed into Or conditions
            // the idea here is instead of queriying for 1 record we query for group of records with Or condition
            string[] itemQueryConditions = new string[count];

            for (int j = 0; j < count; j++)
            {
                IDictionary<string, object> inputItem = userInputData[j].InputData;

                // each SHarepoint list item might have multiple primary key columns
                // we need to prepare query considering each of these primary keys
                string[] itemSpecificPrimaryKeyColQueries = new string[primaryKeysCount];

                for (int i = 0; i < primaryKeysCount; i++)
                {
                    // get the sharepoint specific provider property
                    var field = this.ListPrimaryKeyColumns[i].ProviderSpecificProperty as Field;
                    string fieldName = field.EntityPropertyName;

                    // the type is important here for query purposes, else it may not retrieve correct result
                    itemSpecificPrimaryKeyColQueries[i] = $"<Eq><FieldRef Name=\"{fieldName}\" /><Value Type='{field.FieldTypeKind.ToString()}'>{inputItem[fieldName]}</Value></Eq>";
                }

                // combine each of the primary key column query into group of And
                string combinedQueryForSpecificItem = this.FormatToSharepointQuery(itemSpecificPrimaryKeyColQueries, "And");

                itemQueryConditions[j] = combinedQueryForSpecificItem;
            }

            string whereCondition = this.FormatToSharepointQuery(itemQueryConditions, "Or");
            return whereCondition;
        }

        /// <summary>
        /// Batches the conditions into group of two.
        /// the Or, And Logical operators can only have 2 conditions at once 
        /// and we may have several conditions for query, this methods reduces/combine
        /// compatible to sharepoint query format.
        /// ex: if the input condition is 
        /// <Eq><Field FieldRef="test"><value Type='text'>abcd</value></Eq> --> condit
        /// <Eq><Field FieldRef="test"><value Type='text'>abcd2</value></Eq> --> condi
        /// <Eq><Field FieldRef="test"><value Type='text'>abcd3</value></Eq> --> condi
        /// the output would be <Or> <Or>Condition_1 Condtion_2</Or> condition_3 </Or>
        /// </summary>
        /// <param name="conditions">The conditions.</param>
        /// <param name="logicalOperator">The logical operator which can be And, Or, etc.</param>
        /// <returns>Returns a combiled query based in the input logical operator</returns>
        private string FormatToSharepointQuery(string[] conditions, string logicalOperator)
        {
            // the Or, And Logical operators can only have 2 conditions at once 
            // and we may have several conditions for query, this methods reduces/combines the conditions which can be
            // compatible to sharepoint query format.
            // ex: if the input condition is 
            // <Eq><Field FieldRef="test"><value Type='text'>abcd</value></Eq> --> condition_1
            // <Eq><Field FieldRef="test"><value Type='text'>abcd2</value></Eq> --> condition_2
            // <Eq><Field FieldRef="test"><value Type='text'>abcd3</value></Eq> --> condition_3
            // the output would be <Or> <Or>Condition_1 Condtion_2</Or> condition_3 </Or>

            // collapse them to suit 2 conditions
            List<string> orConditionsList = new List<string>();
            orConditionsList.AddRange(conditions.Where(o => !string.IsNullOrEmpty(o)));

            while (true)
            {
                List<string> newConditionsList = new List<string>();
                int condnCount = orConditionsList.Count();
                for (int orIndex = 0; orIndex < condnCount; orIndex = orIndex + 2)
                {
                    string temp = string.Empty;
                    if (orIndex + 1 < condnCount)
                    {
                        // temp = "<Or>" + orConditionsList[orIndex] + orConditionsList[orIndex + 1] + "</Or>";
                        temp = "<" + logicalOperator + ">" + orConditionsList[orIndex] + orConditionsList[orIndex + 1] + "</" + logicalOperator + ">";
                    }
                    else
                    {
                        temp = orConditionsList[orIndex];
                    }

                    newConditionsList.Add(temp);
                }

                orConditionsList = newConditionsList;

                // if there are no more conditions to club then break from the loop
                if (newConditionsList.Count() == 1)
                {
                    return newConditionsList.FirstOrDefault();
                }
            }
        }
               
        private void ReportProgress(string comment= "", int recordsRead= -1, int recordsUpdated = -1)
        {
            // this.progressLogStream.WriteLine(comment);
            Console.WriteLine(comment);
            AppLogger.Info(comment);

            this.progressReporter.ReportProgress(recordsRead, recordsUpdated);
        }

        private void CompareSourceWithData(List<ListItem> source, List<IWrappedDataItem> updateData)
        {
            int totalUpdateCount = updateData.Count;

            for (int i = 0; i < totalUpdateCount; i++)
            {
                var item = updateData[i];
                SPAction actionType = this.Compare(source, item.InputData);

                item.UpdateMode = actionType;

                if(actionType != SPAction.NoAction)
                {
                    // we enqueue items which needs to be processed into the list action queue
                    // On a different thread these queues will be read and processed.
                    this.spListDataToUpdate.Enqueue(item);
                }
            }
        }

        private SPAction Compare(List<ListItem> source, dynamic target)
        {
            int primaryColCount = this.ListPrimaryKeyColumns.Length;

            var targetDictObj = (IDictionary<string, object>)target;
                        
            var matchObj = source.Where(o =>
              {               
                  bool matchFound = primaryColCount > 0;

                  for (int pColindex = 0; pColindex < primaryColCount && matchFound; pColindex++)
                  {
                      string propertyName = this.ListPrimaryKeyColumns[pColindex].ColumnName;

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
                for (int colIndex = 0; colIndex < this.listColumnsCount; colIndex++)
                {                    
                    string propertyToCompare = this.GetListColumnNameAt(colIndex);

                    needUpdate = !this.CompareValueOrRefType(matchObj, targetDictObj, propertyToCompare);                    
                    if (needUpdate) break;
                }

                if (needUpdate)
                {
                    target.ID = matchObj["ID"];
                    return SPAction.Update;
                }
                else
                    return SPAction.NoAction;
            }
        }

        /// <summary>
        /// Updates the provider specific list item with the values provided by the user
        /// </summary>
        /// <param name="oListItem">The sharepoint specific list item.</param>
        /// <param name="inputItem">The input item from which the value needs to be copied.</param>
        private void UpdateToSPListItem(ListItem oListItem, IDictionary<string, object> inputItem)
        {            
            for (int i = 0; i < this.listColumnsCount; i++)
            {
                string field = this.GetListColumnNameAt(i);
                oListItem[field] = inputItem[field];
            }

            oListItem.Update();
        }

        private bool CompareValueOrRefType(ListItem obj1, IDictionary<string, object> obj2, string propertyName)
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

        private void ExecuteContinousUpdate()
        {
            try
            {
                // this is a continuous thread which monitors the spListDataToUpdate queue for any items in queue
                // if item found for update then it updates it to Sharepoint;
                // updates are performed in predefined batch counts ex: 100 records updated at once.
                using (ClientContext ctx = this.contextHandler.GetContext())
                {
                    if (ctx != null)
                    {
                        var oList = ctx.Web.Lists.GetByTitle(this.ListName);
                        var listCreationinformation = new ListItemCreationInformation();
                        int count = 0;
                        
                        int totalRecordsUpdated = 0;

                        List<IWrappedDataItem> batchCommitList = new List<IWrappedDataItem>();
                        
                        while (!this.allRecordsComparisonComplete || this.spListDataToUpdate.Count() != 0)
                        {
                            if (this.spListDataToUpdate.Count() == 0)
                            {
                                // before going to sleep, check if there are any records to be committed at all, if yes, 
                                // then commit the items 
                                if (count > 0)
                                {
                                    // if there was any error during commit add it to the error list so that it can be 
                                    // retriggered once again in batch of 1, the errors may happen because of duplicates
                                    bool iSsuccess = this.PerformUpdateCommit(ctx, batchCommitList, totalRecordsUpdated);
                                    batchCommitList.Clear();
                                    count = 0;

                                    // reprocess from the beginning after commit as there may be records which could have been added into the queueu
                                    continue;
                                }

                                AppLogger.Info("The Update Thread is sleeping for 5 seconds as there are no records to update");
                                Thread.Sleep(TimeSpan.FromSeconds(5));
                            }

                            IWrappedDataItem nextItem = null;
                           
                            if (this.spListDataToUpdate.TryDequeue(out nextItem))
                            {
                                nextItem.UpdateStatus = UpdateStatus.Inprogress;
                                totalRecordsUpdated++;
                                ListItem oListItem = null;
                                if (nextItem.UpdateMode == SPAction.Insert)
                                {
                                    oListItem = oList.AddItem(listCreationinformation);
                                }
                                else if (nextItem.UpdateMode == SPAction.Update)
                                {
                                    oListItem = oList.GetItemById(nextItem.InputData.ID);
                                }

                                 this.UpdateToSPListItem(oListItem, nextItem.InputData);

                                batchCommitList.Add(nextItem);
                                count++;
                                
                                // commit for every x records ex: 100 records
                                if (count >= this.BatchCommit)
                                {
                                    // if there was any error during commit add it to the error list so that it can be 
                                    // retriggered once again in batch of 1, the errors may happen because of duplicates
                                    bool iSsuccess = this.PerformUpdateCommit(ctx, batchCommitList, totalRecordsUpdated);
                                    batchCommitList.Clear();
                                    count = 0;
                                }
                            }
                        }

                        // if there are some pending items to be updated then update them
                        if (count != 0)
                        {
                            // if there was any error during commit add it to the error list so that it can be 
                            // retriggered once again in batch of 1, the errors may happen because of duplicates
                            bool iSsuccess = this.PerformUpdateCommit(ctx, batchCommitList, totalRecordsUpdated);
                            batchCommitList.Clear();
                        }

                        // if there were any failed Records reprocess them
                        var errorRecords = batchCommitList.Where(o => o.UpdateStatus == UpdateStatus.Failure).ToList();
                        if (errorRecords.Count() > 0)
                        {
                            this.ReAttemptCommitOfErrorRecords(errorRecords);
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

        private void ReAttemptCommitOfErrorRecords(List<IWrappedDataItem> errorRecords)
        {
            try
            {
                // There might have been errors due to duplicate or data issues, we attempt to recover from these
                // errors by commiting in group of one
                using (ClientContext ctx = this.contextHandler.GetContext())
                {
                    if (ctx != null)
                    {
                        int i = 0;
                        var oList = ctx.Web.Lists.GetByTitle(this.ListName);
                        var listCreationinformation = new ListItemCreationInformation();
                        
                        foreach (var nextItem in errorRecords)
                        {
                            ListItem oListItem = null;
                            if (nextItem.UpdateMode == SPAction.Insert)
                            {
                                oListItem = oList.AddItem(listCreationinformation);
                            }
                            else if (nextItem.UpdateMode == SPAction.Update)
                            {
                                oListItem = oList.GetItemById(nextItem.InputData.ID);
                            }

                            this.UpdateToSPListItem(oListItem, nextItem.InputData);                           

                            try
                            {
                                ctx.ExecuteQuery();
                                nextItem.UpdateMode = SPAction.Complete;
                                i++;
                                this.ReportProgress("Records Committed", recordsUpdated: i);                                
                            }
                            catch (Exception ex)
                            {
                                string itemLog = string.Empty;
                                nextItem.UpdateStatus = UpdateStatus.Failure ;
                                itemLog += "\n" + this.GetLogForItem(nextItem.InputData);                                
                                AppLogger.Error(ex, "Error encountered when updating items to Sharepoint, item info: " + itemLog);
                            }
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

        private bool PerformUpdateCommit(ClientContext ctx, List<IWrappedDataItem> batchCommitList,  int totalRecordsUpdated)
        {
            DateTime dt = DateTime.Now;
            try
            {
                ctx.ExecuteQuery();
                this.ReportProgress("Records Committed", recordsUpdated: totalRecordsUpdated);
                batchCommitList.ForEach(o =>
                {
                    o.UpdateMode = SPAction.Complete;
                    o.UpdateStatus = UpdateStatus.Success;
                });

                return true;
            }
            catch (Exception ex)
            {
                string itemLog = string.Empty;
                batchCommitList.ForEach(o => { itemLog += "\n" + this.GetLogForItem(o.InputData); });

                batchCommitList.ForEach(o =>
                {
                    o.UpdateStatus = UpdateStatus.Failure;
                });

                AppLogger.Error(ex, "Error encountered when updating items to Sharepoint, item info: " + itemLog);

                return false;
            }

            // Console.WriteLine($"In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}: at count: {this.TotalCount++}");

            // Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}: In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}, TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");
            // this.ReportProgress($"{DateTime.Now.ToString("HH:mm:ss")}: In Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}, Successful update of batch: {batch++} TimeTaken: {DateTime.Now.Subtract(dt).TotalSeconds} seconds");            
        }


        /// <summary>
        /// Returns the technical column name of a list column at the specified index from the List Columns which the user has selected for Update
        /// </summary>
        /// <param name="index">The index at which the column name has to be retrieved.</param>
        /// <returns>returns the column name at the specified index</returns>
        private string GetListColumnNameAt(int index)
        {
            return this.ListColumnsToUpdate[index].ColumnName;
        }

        private string GetLogForItem(IDictionary<string, object> dataItem)
        {
            string log = string.Empty;
            string propertyName = string.Empty;

            try
            {
                
                for (int i = 0; i < this.listColumnsCount; i++)
                {
                    string fieldName = this.GetListColumnNameAt(i);
                    if (dataItem.ContainsKey(fieldName))
                    {
                        log += fieldName + " = " + dataItem[fieldName] + " , ";
                    }
                    else
                    {
                        log += "ERROR: NO field: " + fieldName + " found in the data Item";
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Unexpected Error encountered when formatting the input data object");
            }

            return log;
        }        

        public enum SPAction
        {
            Blank,

            Insert,

            Delete,

            Update,

            NoAction, 

            Complete
        }
    }
}
