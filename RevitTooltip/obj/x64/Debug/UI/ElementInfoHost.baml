   M S B A M L   `   `   `  ����  RevitTooltip.    (Revit.Addin.RevitTooltip.ElementInfoHost   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      7wnds(http://schemas.microsoft.com/xps/2005/06    5      $��15(�5      &   System.Windows.Media.TextOptions    TextFormattingMode$  Display=�5       � 	Resources 5
       �� MergedDictionaries �� Source�� 5   
    5      �� 5      $J Dpack://application:,,,/RevitTooltip;component/Styles/MetroStyle.xaml?�6%   %    5      .��5      � 
������� � RowDefinitions 5   
   �� 5      $��6* �6   �� 5      $��1* �6   .+�5   
   &   System.Windows.Controls.DataGrid  -   5          Name$ dg��5      $��0q�6      AutoGenerateColumns .     CanUserAddRows .  $
��auto��6�   $
��Auto=�6�      Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn	   
IsReadOnly	 . 
   Header$
 属性名称��6P      Width5  /System.Windows.Controls.DataGridLengthConverter$ 5* 6^      CanUserSort .  1  +System.Windows.Controls.DataGridBoundColumn   Binding 67   �� *Name+   	CellStyle 5      �� 5       �� 
TargetType*  $System.Windows.Controls.DataGridCell"  6     �� Triggers 5      �� 5   "    �� Binding 6.   �� *Name+ �� Value
 备注6G   .a�5   &   ��  �� Property � �� �� Value
 ��  �6C   5   #   5       5"        5#      $
 属性数据��6?   $ 5* 6M    .   6%   �� *Value+5$      5&      � 5'   
   $��1��6   $	��Top=�6   .+�5(      �� -   5        $��30��5(      $��保存��66   $��15(�6C   5)      5*      5+      