using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using SpInteractionHandler;
using SpInteractionHandler.SharepointUpdateHandlers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using GalaSoft.MvvmLight.Messaging;
using SP.SpCommonFun;
using System.ComponentModel;

namespace SharepointBulkUploadTool.ViewModel
{
    public class MainWindowViewModel : ViewModelBase, IReportProgress
    {
        private readonly char[] tabArray = { '\t' };

        IDisplayWindowHandler windowHandler = null;

        IDomainUpdateHandler domainHandler = null;

        private string targetURL = string.Empty;

        private ICommand updateDataToSharepointCommand = null;

        private DateTime updateStartTime;

        private bool isUpdateInProgress;

        private Dictionary<string, ListItemDropDownColumnModel> userSelectedPropNames = new Dictionary<string, ListItemDropDownColumnModel>();
               
        private List<IValidatableItem> allValidatableItems = new List<IValidatableItem>();

        public MainWindowViewModel(IDisplayWindowHandler handler)
        {
            this.windowHandler = handler;

            this.domainHandler = new SharepointHandler(targetURL);
            this.SharepointSiteURLItem = new ViewModelItem<string>(async (sharepointURL) =>
            {
                await this.ConnectToSharepointAndRetrieveListDetails(sharepointURL);
            }, this.windowHandler);            

            this.SelectedList = new ViewModelItem<string>(async (o) =>
            {
                // this is call back method when the value changes
                await this.OnTableSelected(o);
            }, this.windowHandler);

            this.SelectedListPrimaryColumns = new ViewModelItem<ObservableCollection<TableListColumn>>(this.windowHandler,
                (listItems) =>
                {
                    return listItems != null && listItems.Count != 0;
                }
                );

            this.SelectedListPrimaryColumns.Value = PrimaryKeyColumns;

            this.allValidatableItems.Add(this.SharepointSiteURLItem);
            this.allValidatableItems.Add(this.SelectedList);
            this.allValidatableItems.Add(this.SelectedListPrimaryColumns);
        }

        #region properties      

        
        public ObservableCollection<string> AllSharepointListNames { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// Gets or sets the selected table columns. This is a list which would contain all the associated columns 
        /// of the selected table/List
        /// </summary>
        /// <value>
        /// The selected table columns.
        /// </value>
        public ObservableCollection<TableListColumn> SelectedListAllColumns { get; set; } = new ObservableCollection<TableListColumn>();

        public ViewModelItem<ObservableCollection<TableListColumn>> SelectedListPrimaryColumns { get; set; }

        public ViewModelItem<string> SharepointSiteURLItem { get; set; }

        public ViewModelItem<string> SelectedList { get; set; }

        public ICommand UploadToSharepointCommand
        {
            get
            {
                return this.updateDataToSharepointCommand ?? (
                    this.updateDataToSharepointCommand = new RelayCommand(() =>
                    {
                        if (this.IsUpdateInProgress)
                            return;

                        Task t = Task.Run(() =>
                          {
                              try
                              {
                                  // indicate that the work has now started
                                  this.IsUpdateInProgress = true;
                                  bool isValid = true;

                                  this.allValidatableItems.ForEach(o =>
                                  {
                                      isValid &= o.Validate();
                                  });

                                  if (isValid == false)
                                      return;

                                  this.updateStartTime = DateTime.Now;

                                  // start a timer which can indicate the total time passed
                                  using (Timer totalTimeRun = new Timer(this.UpdateTotalTimeElapsed, null, 0, 1000))
                                  {
                                      TableListColumn[] headerCols = this.userSelectedPropNames
                                                                       .Values.Where(o => o != null && o.SelectedColumn != null).Select(o => o.SelectedColumn).ToArray();
                                      TableListColumn[] primaryCols = this.PrimaryKeyColumns.ToArray();

                                      this.NumberOfRecordsInputByUser = this.FormattedTextDataDyn.Count();
                                      this.RaisePropertyChanged(nameof(this.NumberOfRecordsInputByUser));

                                      this.MapFromDefaultNameToUserSelectedProperty(this.FormattedTextDataDyn, headerCols);

                                      this.domainHandler.UpdateListToSource(
                                          this.SelectedList.Value,
                                          this.FormattedTextDataDyn,
                                          headerCols,
                                          primaryCols,
                                          this
                                          );
                                  }
                              }
                              catch (Exception ex)
                              {
                                  AppLogger.Error(ex, "Unexpected error when invoking the Upload Operation");
                                  this.windowHandler.ShowErrorMessage("Unexpected Error: " + ex.Message);
                                  // throw;
                              }
                              finally
                              {
                                  // indicate that the work is complete.
                                  this.IsUpdateInProgress = false;
                              }

                          });
                    },
                    () =>
                    {
                        return this.IsUpdateInProgress == false;
                    }));
            }
        }

        public ObservableCollection<TableListColumn> PrimaryKeyColumns { get; set; } = new ObservableCollection<TableListColumn>();

        public List<IWrappedDataItem> FormattedTextDataDyn { get; set; } = new List<IWrappedDataItem>();

        #endregion
        
        public ListItemDropDownColumnModel GetColumnBindingItem(int colIndex)
        {
            if (this.SelectedListItem != null)
            {
                string columnName = this.GetDefaultColName(colIndex);

                ListItemDropDownColumnModel itemToReturn = null;
                if (this.userSelectedPropNames.ContainsKey(columnName) == false)
                {
                    itemToReturn = new ListItemDropDownColumnModel();
                    itemToReturn.AllColumns = this.SelectedListItem.ListColumns;
                    this.userSelectedPropNames[columnName] = itemToReturn;
                }

                return this.userSelectedPropNames[columnName];
            }

            return null;
        }

        public TableListItem SelectedListItem { get; set; }

        private string tsvData;

        /// <summary>
        /// Gets or sets the Tab Seperated data text which the user copies from clipboard
        /// </summary>
        /// <value>
        /// The TSV data text.
        /// </value>
        public string TsvDataText
        {
            get { return tsvData; }
            set { tsvData = value; this.FormatText(); }
        }

        private void FormatText()
        {
            this.userSelectedPropNames.Clear();
            this.FormattedTextDataDyn.Clear();
            var inputDataObjects = this.FormattedTextDataDyn;
            string[] allRows = this.tsvData.Split('\n');
            int rowCount = allRows.Count();

            string[] headerCols = allRows[0].Split(tabArray, StringSplitOptions.RemoveEmptyEntries).Select(o => o.Trim()).ToArray();

            for (int i = 1; i < rowCount; i++)
            {
                string[] inputData = allRows[i].Split(tabArray).Select(o => o.Trim()).ToArray();

                int colCount = inputData.Count();
                dynamic dyn = new ExpandoObject();
                var dic = (IDictionary<string, object>)dyn;

                for (int col = 0; col < colCount; col++)
                {
                    string colName = this.GetDefaultColName(col);
                    dic[colName] = inputData[col];
                }

                inputDataObjects.Add(new WrappedListItem() { InputData = dic });
            }

            this.NumberOfRecordsInputByUser = inputDataObjects.Count();
            this.windowHandler.UpdateGridColumns(inputDataObjects);
        }

        private void MapFromDefaultNameToUserSelectedProperty(List<IWrappedDataItem> inputSourceDataList, TableListColumn[] headerCols)
        {
            // By default the column names are given name col1, col2, etc; 
            // These are auto-assigned by system; user would then map each of these column to a meaning full list name
            // in this method we add additional property to the dynamic object by mapping col1 to the corresponding header col
            int totalFields = headerCols.Length;

            // prepare a default property array which maps 1:1 system generated name to user selected property name
            string[] systemGeneratedColNames = new string[totalFields];
            string[] userSelectedPropNames = new string[totalFields];

            for (int i = 0; i < totalFields; i++)
            {
                systemGeneratedColNames[i] = this.GetDefaultColName(i);
                userSelectedPropNames[i] = headerCols[i].ColumnName;
            }

            int totalInputData = inputSourceDataList.Count;
            for (int inpIndex = 0; inpIndex < totalInputData; inpIndex++)
            {
                dynamic sourceObject = inputSourceDataList[inpIndex].InputData;

                IWrappedDataItem item = new WrappedListItem();

                item.InputData = sourceObject;

                var dict = (IDictionary<string, object>)sourceObject;

                for (int i = 0; i < totalFields; i++)
                {
                    var headerColProp = headerCols[i];
                    string systemGeneratedColName = systemGeneratedColNames[i];
                    object value = null;

                    if (dict.ContainsKey(systemGeneratedColName))
                    {
                        value = dict[systemGeneratedColName];
                    }


                    if (headerColProp.ColumnType == typeof(DateTime))
                    {
                        DateTime objectVal;
                        if (DateTime.TryParse(value.ToString(), out objectVal) == false)
                        {
                            dict["HasError"] = true;
                            dict["ErrorDescription"] = "Unable to convert the column: " + headerColProp.ColumnDisplayName + " to type DateTime";
                        }

                        value = objectVal;
                    }

                    dict[userSelectedPropNames[i]] = value;
                }
            }
        }

        /// <summary>
        /// Gets the default name of the col. This is an autogenerated name for the property. 
        /// User enters a tab seperated data in the Input data tab. Each of these tab seperated data needs to be convereted
        /// into a dynamic object with a default property name. This Method returns a property name based on the index
        /// </summary>
        /// <param name="colIndex">Index of the col.</param>
        /// <returns></returns>
        public string GetDefaultColName(int colIndex)
        {
            return "__1COL_" + colIndex;
        }

        private void InvokeOnUIThread(Action methodToInvoke)
        {
            Delegate method = methodToInvoke;
            var result = System.Windows.Application.Current.Dispatcher.Invoke(method, null);
        }

        public int NumberOfRecordsRead { get; set; }

        private int numberOfRecordsInputByUser;

        public int NumberOfRecordsInputByUser
        {
            get
            {
                return this.numberOfRecordsInputByUser;
            }
            set
            {
                this.numberOfRecordsInputByUser = value;
                this.RaisePropertyChanged();
            }
        }

        public List<SpInteractionHandler.UpdateStatus> AllStatus
        {
            get
            {
                return Enum.GetValues(typeof(SpInteractionHandler.UpdateStatus)).Cast<SpInteractionHandler.UpdateStatus>().ToList();
            }
        }

        public int NumberOfRecordsUpdated { get; set; }

        public int NumberOfRecordsInError { get; set; }

        public string TotalTimeElapsed { get; set; }

        public void ReportProgress(int numberOfRecordsRead = -1, int numberOfRecordsUpdated = -1)
        {
            this.NumberOfRecordsRead = numberOfRecordsRead == -1 ? this.NumberOfRecordsRead : numberOfRecordsRead;
            this.NumberOfRecordsUpdated = numberOfRecordsUpdated == -1 ? this.NumberOfRecordsUpdated : numberOfRecordsUpdated;

            base.RaisePropertyChanged(nameof(NumberOfRecordsUpdated));
            base.RaisePropertyChanged(nameof(NumberOfRecordsRead));
        }

        private void UpdateTotalTimeElapsed(object state)
        {
            this.TotalTimeElapsed = DateTime.Now.Subtract(updateStartTime).ToString(@"dd\.hh\:mm\:ss");

            this.RaisePropertyChanged(nameof(this.TotalTimeElapsed));
        }

        /// <summary>
        /// Determines whether the sharepoint list can be updated, validates for all the user inputs
        /// </summary>
        /// <returns>Returns true if the update can be started</returns>
        private bool CanStartUpdate()
        {
            return true;
        }

        public bool IsUpdateInProgress
        {
            get { return isUpdateInProgress; }
            set
            {
                isUpdateInProgress = value;
                this.RaisePropertyChanged();
            }
        }

        private Task ConnectToSharepointAndRetrieveListDetails(string sharepointURL)
        {
            return Task.Run(() =>
             {
                 try
                 {
                     bool isSuccess = false;

                    // activate the spinner
                    this.InvokeOnUIThread(() =>
                     {
                        //this.SharepointSiteURLItem.SetItemStatus(FieldStatus.InProgress);

                        this.AllSharepointListNames.Clear();

                         this.domainHandler = new SharepointHandler(sharepointURL);
                         isSuccess = this.domainHandler.TestConnectivity();
                     });


                     if (isSuccess == false)
                     {
                         this.SharepointSiteURLItem.SetValidity(isValid: false, message: "The entered URL is invalid");
                         return;
                     }

                     var allItems = this.domainHandler.GetAllUpdateableItems();

                     this.InvokeOnUIThread(() =>
                     {
                         allItems.ForEach(o => this.AllSharepointListNames.Add(o));

                        //this.SharepointSiteURLItem.SetItemStatus(FieldStatus.Success);
                    });

                     this.SharepointSiteURLItem.SetValidity(isValid: true);
                 }
                 catch (Exception ex)
                 {
                     AppLogger.Error(ex, "Exception encountered when validating sharepoint details");
                     this.windowHandler.ShowErrorMessage("Unable to validate Sharepoint Details.\nException Details: " + ex.ToString(), "Critical Error");
                 }
             });
        }

        private Task OnTableSelected(string selectedlistName)
        {
            Console.WriteLine("Thread: " + Thread.CurrentThread.ManagedThreadId);
            return Task.Run(() =>
            {
                this.InvokeOnUIThread(() =>
                {
                    // add all the columns assoicated with the selected list into an observalbe collection for user display
                    this.SelectedListAllColumns.Clear();
                });

                if (!string.IsNullOrEmpty(selectedlistName))
                {
                    this.SelectedListItem = this.domainHandler.GetListItem(selectedlistName);

                    // this needs to run on UI Thread as it invovles modifying the UI bound colleciton
                    this.InvokeOnUIThread(() =>
                    {
                        this.SelectedListItem.ListColumns.ForEach(o =>
                            this.SelectedListAllColumns.Add(o)
                        );
                    });
                }
            }
              );

        }
    }

    public class ListItemDropDownColumnModel
    {
        public TableListColumn SelectedColumn { get; set; }

        public List<TableListColumn> AllColumns { get; set;}
    }

    public class WrappedListItem : ViewModelBase, IWrappedDataItem
    {
        private dynamic inputdata;
        private SPListUpdateHandler.SPAction updateMode;

        private SpInteractionHandler.UpdateStatus updateStatus;
        public dynamic InputData
        {
            get
            {
                return this.inputdata;
            }

            set
            {
                this.inputdata = value;

                this.RaisePropertyChanged();
            }
        }

        public SPListUpdateHandler.SPAction UpdateMode
        {
            get
            {
                return this.updateMode;
            }

            set
            {
                this.updateMode = value;

                this.RaisePropertyChanged();
            }
        }

        public SpInteractionHandler.UpdateStatus UpdateStatus
        {
            get
            {
                return this.updateStatus;
            }

            set
            {
                this.updateStatus = value;

                this.RaisePropertyChanged();
            }
        }
    }
}


