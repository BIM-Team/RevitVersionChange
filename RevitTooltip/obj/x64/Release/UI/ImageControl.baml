   M S B A M L   `   `   `  ���� G  CRevitTooltip, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullI)clr-namespace:Revit.Addin.RevitTooltip.UIRevit.Addin.RevitTooltip.UI  .    (Revit.Addin.RevitTooltip.UI.ImageControl   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     ;local)clr-namespace:Revit.Addin.RevitTooltip.UI     $
��auto��  � 	Resources   �� MergedDictionaries �� Source��  �� $J Dpack://application:,,,/RevitTooltip;component/Styles/MetroStyle.xaml?�%    .��� 
�������$y�Visible=� � RowDefinitions �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ���  � ColumnDefinitions �� $
��auto ��� $��* �.+��� -    ��Name$ label2��$��类型��$
��Left=�$��0q�$��1��$	��Top=�� $��1�� �� $
��auto ��� $��* �.+��� -    ��Name$ comboBox��$��0��$
��Left=�$��0q�$	��Top=�$	��120�� �� SelectedValuePath$ Signal�� �� DisplayMemberPath$ CurrentFile���� -   $ detail��$��测斜详情��$��1��$��Center=�$y�Hidden=��� -   $ label��$��开始时间��$��Right=�$��0q�$��Center=�$��1��� $��1��$��1�� �� $
��auto ��� $��* �.+��� -   	 ��Name$	 startBox��$
��Left=�$��23��$��0q�
 �� TextWrapping$

 Wrap=�$��Center=�$	��120�� �� 
IsReadOnly . �� -   $ startCancel��$��重置��$��1��$��0q�$��Center=�&   System.Windows.Controls.Calendar  -     Name$ 	startTime��$��Center=�$	��Top=�$��0q�$��2�� K� ZIndex$ 10��$y�Hidden=�$��5��$��2���� -   $ label1��$��结束时间��$��3��$��Right=�$��0q�$��Center=�� $��1��$��3�� �� $
��auto ��� $��* �.+��� -	   $	 endBox��$
��Left=�$��23��$��0q�$

 Wrap=�$��Center=�$	��120�� . �� -
   $ 	endCancel��$��重置��$��1��$��0q�$��Center=�  -   $ endTime��$��Center=�$��0q�$	��Top=�$��2��$ 10��$y�Hidden=�$��4��$��4��&   System.Windows.Controls.DataGrid  -     Name$ dataGrid��$
��Left=�$��0,0,0,0q�$	��Top=�   
IsReadOnly . $��5��$��2��   AutoGenerateColumns .     SelectionUnit$ FullRow=�   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Id��   CanUserSort .     
Visibility$ Hidden=�1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *Id+  $ 	实体名��   Width5  /System.Windows.Controls.DataGridLengthConverter$ 5*  �� *
EntityName+   	CellStyle ��  �� 
TargetType*  $System.Windows.Controls.DataGridCell"   �� Triggers ��  �� Binding �� *ErrMsg+ �� Value 	异常一.a���  �� Property � �� �� Value
 �  ����  �� *ErrMsg+ 两种异常.a���  � �� �	#4C11111C��  � ��
 �  ����  �� *ErrMsg+ 	异常二.a���  � ��
 ��  �  $ 错误��$ 5*  �� *ErrMsg+ �� "   ��  �� *ErrMsg+ 	异常一.a���  � ��
 �  ����  �� *ErrMsg+ 两种异常.a���  � �� �	#4C11111C��  � ��
 �  ����  �� *ErrMsg+ 	异常二.a���  � ��
 ��  �