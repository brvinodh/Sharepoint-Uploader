   M S B A M L   `   `   `  ���� U  QMahApps.Metro, Version=1.5.0.23, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  S OSharepointBulkUploadTool, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullC&clr-namespace:SharepointBulkUploadToolSharepointBulkUploadTool Y1clr-namespace:SharepointBulkUploadTool.Converters#SharepointBulkUploadTool.Converters O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem )   #SharepointBulkUploadTool.MainWindow   HControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro  \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     8local&clr-namespace:SharepointBulkUploadTool     MlocalConverters1clr-namespace:SharepointBulkUploadTool.Converters     Z VFontAwesome.WPF, Version=4.7.0.37774, Culture=neutral, PublicKeyToken=0758b07a11a4f466-fa$http://schemas.fontawesome.io/icons/  3System&clr-namespace:System;assembly=mscorlib    � Title$!  Sharepoint Bulk Upload Tool��$
��1000��$
��1000�� � WindowState$ 	Maximized=�$��2q�
���   � � 	Resources G  ASharepointBulkUploadTool.Converters.VMColorToWindowColorConverter   vmColorConverter  	BoolToVis��%   &	        &	        �� .��� $��10q� � RowDefinitions �� $��* ��� $��30 �.+��� .@��� -   $��ListInfo��$��1. List Information��.��� $��50,50, 0, 0q�$
��Auto��$
��Left=� � ColumnDefinitions �� $
��0.3* ��� $
��0.5* � �� $��35 ��� $��35 ��� $	��500 �.+��� $��0��$��0��$��Enter Sharepoint Site: ��$��Center=�$��Right=�$��0,0,10,0q��� $��0��$��1�� �� Orientation$ 
Horizontal=� � DataContext �� *SharepointSiteURLItem+.+��� $��30��$	��500��$
��Left=�$��Center=����� *Value+"  FontAwesome.WPF.ImageAwesome  $��Center=�$��Center=�$	��5,0q�$��30��$��25��   Icon �� *ItemStatusIcon+   Spin �� *IsInProgress+	   
Foreground	 �� *	IconColor+
 �� 	Converter#
 [  �� $��1��$��0��$��Select List: ��$��Center=�$��Right=�$��0,0,10,0q��� $��1��$��1��$ 
Horizontal=� �� *SelectedList+.+��� -   $��AllLists��$��30��$	��500��$
��Left=�$��Center=����� *$"DataContext.AllSharepointListNames+ �� RelativeSource ��  �� Mode$ FindAncestor=� �� AncestorType$ WindowL� �� SelectedItem �� *Value+  $��Center=�$��Center=�$	��5,0q�$��30��$��25�� �� *ItemStatusIcon+ �� *IsInProgress+	 �� *	IconColor+#
 [  �� $��3��$��0��$'��!Select Primary Key(s) for Update:��$	��Top=�$��Right=�$��0,0,10,0q��� $��3��$��1��$	��Top=� �� *SelectedListPrimaryColumns+.+��� $��1��$��1��$ 
Horizontal=�.+��� -   $��PrimaryKeyColsListBox��$��3��$��1�� �� SelectionMode$ Multiple=�$	��200��$	��500��$
��Left=� �� DisplayMemberPath$ ColumnDisplayName��$��Center=�
�������$��1q����� *$"DataContext.SelectedListAllColumns+ �� $ FindAncestor=�$ WindowL�  $��Center=�$��Center=�$	��5,0q�$��30��$��25�� �� *ItemStatusIcon+ �� *IsInProgress+	 �� *	IconColor+#
 [  �� $��Selected Primary Key(s)��y���  �� ElementName$ selectedPrimaryKeyItems�� �� Path$ Items.Count�#
 [ �� -   $��selectedPrimaryKeyItems������ *DataContext.PrimaryKeyColumns+ �� $ FindAncestor=�$ WindowL����� .R��� $
��Bold"����� *ColumnDisplayName+�� -   $��tsvData��$��2. Update Data��.���� -   $��CopyPasteTextBox��$��1�� �� 
AcceptsTab .  �� AcceptsReturn .  �� TextWrapping$ NoWrap=����� *TsvDataText+.���� -   $��formattedDataOutput��$��3. Formatted Data��.���  �� $��* ��� $��35 �.+�&   System.Windows.Controls.DataGrid  -   $��dataGrid��   AutoGenerateColumns .    CanUserAddRows .     
IsReadOnly .  �� ���� ]� $��1��.+��� $
��Left=�$ 
Horizontal=�.+��� $��
Filter by:��$��Center=��� -	   $��25��$��5,0,0,0q����� *	AllStatus+�� $��Center=�$��10,0,0,0q��� $��Filtered Count: �� �� $
��Bold"����� $ dataGrid��$ Items.Count��� $
��Left=�$��Center=�$��10,0,0,0q��� $�Total Input Records: �� �� $
��Bold"����� *NumberOfRecordsInputByUser+ �� $�Total Records Read: �� �� .��� ���� *NumberOfRecordsRead+ �� $!�Number of Records Updated: �� �� $
��Bold"����� *NumberOfRecordsUpdated+ �� �� $��Records in Error: �� �� .��� ���� *NumberOfRecordsInError+�� $��Right=�$��Center=�.+��� -
   $��totalTimeTaken��$��Right=�y��� *IsUpdateInProgress+�� $��Total Time taken: �� �� $
��Bold"����� *TotalTimeElapsed+� $��Right=� � IsIndeterminate . $��Right=�$��3��$��Center=����� $ totalTimeTaken��$ ActualWidth�y��� *IsUpdateInProgress+#
 [ ]� $��1��.+��� $ 
Horizontal=�$��Right=�$��Right=�.+��� $��Start Update��$	��100��$��0,0,10,0q����� *UploadToSharepointCommand+