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
��Left=�6-   $��0q�6I   $��1��6U   $	��Top=�6f   � 5   
   $��1��6    5       �� 5!      $
��auto �6#   62   �� 5"      $��* �6#   6/   .+�5$   
   �� -   5         ��Name$ comboBox��5$      $��0��6%   $
��Left=�65   $��0q�6P   $	��Top=�6\   $	��120��6t    �� SelectedValuePath$ Signal��6�    �� DisplayMemberPath$ CurrentFile��6�   �� 5%      -   5        $ detail��5%      $��测斜详情��6$   $��1��63   $��Center=�6`   $y�Hidden=�6}   5&      �� 5'   
   -   5        $ label��5'      $��开始时间��6   $��Right=�6.   $��0q�6J   $��Center=�6U   $��1��6p   � 5(   
   $��1��6   $��1��6    5)      �� 5*      $
��auto �6#   62   �� 5+      $��* �6#   6/   .+�5-      �� -   5        	 ��Name$	 startBox��5-      $
��Left=�6)   $��23��6D   $��0q�6P   
 �� TextWrapping$

 Wrap=�6[   $��Center=�6p   $	��120��6�    �� 
IsReadOnly . 6�   �� 5.      -   5        $ startCancel��5.      $��重置��6)   $��1��66   $��0q�6F   $��Center=�6Q   6n   5/      &   System.Windows.Controls.Calendar  50   
   -   5          Name$ 	startTime��50      $��Center=�6'   $	��Top=�6E   $��0q�6]   $��2��6h    K� ZIndex$ 10��6|   $y�Hidden=�6�   $��5��6�   $��2��6�   �� 51   
   -   5        $ label1��51      $��结束时间��6    $��3��6/   $��Right=�6<   $��0q�6X   $��Center=�6c   � 52   
   $��1��6   $��3��6    53      �� 54      $
��auto �6#   62   �� 55      $��* �6#   6/   .+�57      �� -	   5        $	 endBox��57      $
��Left=�6(   $��23��6C   $��0q�6O   $

 Wrap=�6Z   $��Center=�6n   $	��120��6�    . 6�   �� 58      -
   5        $ 	endCancel��58      $��重置��6'   $��1��64   $��0q�6D   $��Center=�6O   6l   59        5:   
   -   5        $ endTime��5:      $��Center=�6$   $��0q�6A   $	��Top=�6L   $��2��6d   $ 10��6y   $y�Hidden=�6�   $��4��6�   $��4��6�   &   System.Windows.Controls.DataGrid  5<   
   -   5          Name$ dataGrid��5<      $
��Left=�6%   $��0,0,0,0q�6@   $	��Top=�6R      
IsReadOnly . $��5��6�   $��2��6�      AutoGenerateColumns .     SelectionUnit$ FullRow=�6�      Columns 5=      0  *System.Windows.Controls.DataGridTextColumn  5>      ,  &System.Windows.Controls.DataGridColumn   Header$ Id��6=      CanUserSort .     
Visibility$ Hidden=�6]   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6%   �� *Id+6]     5?      $ 	实体名��6D      Width5  /System.Windows.Controls.DataGridLengthConverter$ 5* 6Q    6%   �� *
EntityName+   	CellStyle 5@      �� 5A       �� 
TargetType*  $System.Windows.Controls.DataGridCell"  6     �� Triggers 5B      �� 5C   "    �� Binding 6.   �� *ErrMsg+ �� Value 	异常一6I   .a�5D   &   ��  �� Property � �� �� Value
 �  ��6C   5E   #   �� 5F   "    6.   �� *ErrMsg+ 两种异常6I   .a�5G   &   ��  � �� �	#4C11111C6C   �� 5H   &    � ��
 �  ��6C   5I   #   �� 5J   "    6.   �� *ErrMsg+ 	异常二6I   .a�5K   &   ��  � ��
 ��  �6C   5L   #   5N      5P        5Q      $ 错误��6'   $ 5* 63    6>   �� *ErrMsg+ 5R      �� 5S      "  6     5T      �� 5U   "    6.   �� *ErrMsg+ 	异常一6I   .a�5V   &   ��  � ��
 �  ��6C   5W   #   �� 5X   "    6.   �� *ErrMsg+ 两种异常6I   .a�5Y   &   ��  � �� �	#4C11111C6C   �� 5Z   &    � ��
 �  ��6C   5[   #   �� 5\   "    6.   �� *ErrMsg+ 	异常二6I   .a�5]   &   ��  � ��
 ��  �6C   5^   #   5`      5b      5d      5f      5g      