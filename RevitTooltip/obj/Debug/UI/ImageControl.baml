   M S B A M L   `   `   `  ����G  CRevitTooltip, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullI)clr-namespace:Revit.Addin.RevitTooltip.UIRevit.Addin.RevitTooltip.UI  5      .    (Revit.Addin.RevitTooltip.UI.ImageControl   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      ;local)clr-namespace:Revit.Addin.RevitTooltip.UI     5      $
��auto��5	        � 	Resources  5
       �� MergedDictionaries �� Source�� 5   
    5      �� 5      $J Dpack://application:,,,/RevitTooltip;component/Styles/MetroStyle.xaml?�6%   %    5      .��5      � 
�������$y�Visible=�6    � RowDefinitions 5   
   �� 5      $
��auto �6   6,   �� 5      $
��auto �6   6,   �� 5      $
��auto �6   6,   �� 5      $
��auto �6   6,   �� 5      $
��auto �6   6,   �� 5      6    � ColumnDefinitions 5   
   �� 5      $
��auto �6   6.   �� 5      $��* �6   6+   .+�5   
   �� -   5         ��Name$ label2��5      $��类型��6    $
��Left=�6-   $��0q�6I   $��1��6U   $	��Top=�6f   �� 5   
   -   5         ��Name$ comboBox��5      $��1��6%   $
��Left=�65   $��0q�6P   $	��Top=�6\   $	��120��6t    �� Text 6�   �� *	TableDesc+6�   �� 5    
   -   5        $ label��5       $��开始时间��6   $��Right=�6.   $��0q�6J   $��Center=�6U   $��1��6p   �� 5!   
   -   5         ��Name$ startBox��5!      $
��Left=�6%   $��23��6@   $��0q�6L   	 �� TextWrapping$
	 Wrap=�6W   $��Center=�6l   $	��120��6�   $��1��6�   
 �� 
IsReadOnly
 . $��1��6�   &   System.Windows.Controls.Calendar  5"   
   -   5          Name$ 	startTime��5"      $��Center=�6'   $	��Top=�6E   $��0q�6]   $��2��6h    K� ZIndex$ 10��6|   $y�Hidden=�6�   $��5��6�   $��2��6�   �� 5#   
   -   5        $ label1��5#      $��结束时间��6    $��3��6/   $��Right=�6<   $��0q�6X   $��Center=�6c   �� 5$   
   -   5        $ endBox��5$      $��1��6"   $��3��62   $
��Left=�6@   $��23��6[   $��0q�6g   $
	 Wrap=�6r   $��Center=�6�   $	��120��6�   
 . 6�     5%   
   -   5        $ endTime��5%      $��Center=�6$   $��0q�6A   $	��Top=�6L   $��2��6d   $ 10��6y   $y�Hidden=�6�   $��4��6�   $��4��6�   &   System.Windows.Controls.DataGrid  5'   
   -	   5          Name$ dataGrid��5'      $
��Left=�6%   $��0,0,0,0q�6@   $	��Top=�6R      
IsReadOnly . $��5��6�   $��2��6�     5(      )  #System.Windows.Controls.DataGridRow �� 
TargetType   err �� Triggers �� Binding �� Self �� RelativeSource �� Path �� Value �� Property �� Value��%�  &	        �� 5*      "  6    5+      �� 5,       6&   �� # Z $
 Item�
 error16r   .a�5-      ��  � ��
 �����6;   5.      �� 5/       6&   �� # Z $ Value�
 error26s   .a�50      ��  � ��
 �����6;   51      �� 52       6&   �� # Z $ Value�
 error36s   .a�53      ��  � ��
 �����6;   54      56      57         Columns 58      0  *System.Windows.Controls.DataGridTextColumn  59      ,  &System.Windows.Controls.DataGridColumn   Header$ Id��6=   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6%   �� *Id+6=     5:      $ 错误��6>    6%   �� *Err+6>     5;      $ 	实体名��6D    6%   �� *
EntityName+6D   5=      5?      5@      