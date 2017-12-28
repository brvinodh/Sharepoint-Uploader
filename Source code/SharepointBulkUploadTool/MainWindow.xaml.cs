using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using SharepointBulkUploadTool.ViewModel;
using SpInteractionHandler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SharepointBulkUploadTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow, IDisplayWindowHandler
    {
        MainWindowViewModel model = null;
        public MainWindow()
        {
            InitializeComponent();
            this.model = new MainWindowViewModel(this);
            this.DataContext = model;
        }

        public void ShowErrorMessage(string message, string header = "")
        {
            Action a = async () =>
            {
                await this.ShowMessageAsync(message, header);
            };

            Dispatcher.BeginInvoke(a);
        }

        public async void ShowMessage(string message, string header = "")
        {
            await this.ShowMessageAsync(message, header);
        }

        public void UpdateGridColumns(object source)
        {
            this.dataGrid.DataContext = null;
            this.dataGrid.Columns.Clear();
            this.dataGrid.ItemsSource = null;
            this.dataGrid.Items.Clear();
            //this.dataGrid.DataContext = (source as DataTable).DefaultView;
            //return;


            // Since there is no guarantee that all the ExpandoObjects have the 
            // same set of properties, get the complete list of distinct property names
            // - this represents the list of columns
            this.dataGrid.AutoGenerateColumns = false;
            //var rows = (source as List<dynamic>).OfType<IDictionary<string, object>>();
            var firstItem = (source as List<IWrappedDataItem>).FirstOrDefault();

            if(firstItem == null)
            {                
                return;
            }

            this.dataGrid.ItemsSource = source as IEnumerable<IWrappedDataItem>;

            //var rows = dataGrid.ItemsSource.OfType<IDictionary<string, object>>();
            //var columns = rows.SelectMany(d => d.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            string inputDataPropName = nameof(IWrappedDataItem.InputData);
            var columns = ((IDictionary<string, object>)firstItem.InputData).Keys.Select(o => inputDataPropName + "." + o).ToArray();
            

            int colCount = columns.Count();

            for (int i = 0; i < colCount; i++)
            {
                ListItemDropDownColumnModel dropDownContextItem = this.model.GetColumnBindingItem(i);
                var cboBox = new ComboBox()
                {
                    DataContext = dropDownContextItem,
                    ItemsSource = dropDownContextItem == null ? null : dropDownContextItem.AllColumns,
                    DisplayMemberPath = nameof(TableListColumn.ColumnDisplayName)
                };
                
                Binding b = new Binding(nameof(ListItemDropDownColumnModel.SelectedColumn));
                //b.Source = dropDownContextItem.SelectedColumn;

                cboBox.SetBinding(ComboBox.SelectedItemProperty, b);                

                // now set up a column and binding for each property
                var column = new DataGridTextColumn
                {
                    Header = cboBox,
                    Binding = new Binding(columns[i])
                };

                dataGrid.Columns.Add(column);
            }

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Update Mode",
                Binding = new Binding(nameof(IWrappedDataItem.UpdateMode))
            });

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Update Status",
                Binding = new Binding(nameof(IWrappedDataItem.UpdateStatus)),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            });

            // this.SetListViewSummary(source);

            //CollectionView view = new CollectionView(this.dataGrid.ItemsSource);
            //this.lvSummary.ItemsSource = source as IEnumerable<IWrappedDataItem>;
            //var x = CollectionViewSource.GetDefaultView(this.lvSummary.ItemsSource);
            //this.ActiveLiveGrouping(x, new List<string>() { "UpdateStatus"});
        }


        private void ActiveLiveGrouping(ICollectionView collectionView, IList<string> involvedProperties)
        {
            var collectionViewLiveShaping = collectionView as ICollectionViewLiveShaping;
            if (collectionViewLiveShaping == null) return;

            collectionViewLiveShaping.IsLiveGrouping = true;
            if (collectionViewLiveShaping.CanChangeLiveGrouping)
            {
                foreach (string propName in involvedProperties)
                    collectionViewLiveShaping.LiveGroupingProperties.Add(propName);
                collectionViewLiveShaping.IsLiveGrouping = true;
            }

            collectionView.Refresh();
        }

        private void PrimaryKeyColList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.model.PrimaryKeyColumns.Clear();
            foreach (var item in this.PrimaryKeyColsListBox.SelectedItems)
            {
                this.model.PrimaryKeyColumns.Add(item as TableListColumn);
            }
        }

        private void FilterValue_Changed(object sender, SelectionChangedEventArgs e)
        {
            ListCollectionView filterView = CollectionViewSource.GetDefaultView(this.dataGrid.ItemsSource) as ListCollectionView;
            filterView.IsLiveFiltering = true;

           
            
            filterView.LiveFilteringProperties.Add(nameof(WrappedListItem.UpdateStatus));
            var currentStatus = (e.Source as ComboBox).SelectedItem;
            //CityName = lstCity.SelectedItem.ToString();
            
            Func<object, bool> action = (object d) =>
                {
                    bool result = false;
                    WrappedListItem std = d as WrappedListItem;
                    if (std.UpdateStatus == (SpInteractionHandler.UpdateStatus)currentStatus)
                    {
                        result = true;
                    }                    
                    return result;
                };

            filterView.Filter = new Predicate<object>(action);
            
            filterView.Refresh();
        }
    }
}
