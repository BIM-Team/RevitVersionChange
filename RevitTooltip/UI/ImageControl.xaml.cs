﻿using Autodesk.Revit.UI;
using Revit.Addin.RevitTooltip.Dto;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Globalization;

namespace Revit.Addin.RevitTooltip.UI
{
    /// <summary>
    /// ImageControl.xaml 的交互逻辑
    /// </summary>
    public partial class ImageControl : Page,IDockablePaneProvider
    {
        private Guid m_guid = new Guid("502805E8-5698-4428-A15B-0E8BADE393E0");
        public Guid Id
        {
            get { return m_guid; }
        }
        private ImageControl()
        {
            InitializeComponent();
           // List<ExcelTable> allTypes = App.Instance.Sqlite.SelectDrawTypes();
            //this.comboBox.ItemsSource = allTypes;
        }
        private static ImageControl _image;
        public static ImageControl Instance(){
            if (_image == null) {
                _image = new ImageControl();
            }
            return _image;
            } 
        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CEntityName selectedItem = this.dataGrid.SelectedItem as CEntityName;
            DateTime? start = this.startTime.SelectedDate as DateTime?;
            DateTime? end = this.startTime.SelectedDate as DateTime?;
            NewImageForm.Instance().EntityData= App.Instance.Sqlite.SelectDrawEntityData(selectedItem.EntityName, start, end);
            if (!NewImageForm.Instance().Visible) {
            NewImageForm.Instance().Show();
            }
        }
        public void setDataSource(IEnumerable itemsSource) {
            dataGrid.ItemsSource = itemsSource;
        }

        private void startBox_GotFocus(object sender, RoutedEventArgs e)
        {
            startTime.Visibility = Visibility.Visible;
        }

        private void endBox_GotFocus(object sender, RoutedEventArgs e)
        {
            endTime.Visibility = Visibility.Visible;
        }

        private void startTime_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            startBox.Text = Convert.ToDateTime(startTime.SelectedDate).ToString("yyyy/MM/dd");
            startTime.Visibility = Visibility.Hidden;
        }

        private void endTime_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            endBox.Text = Convert.ToDateTime(endTime.SelectedDate).ToString("yyyy/MM/dd");
            endTime.Visibility = Visibility.Hidden;
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this as FrameworkElement;
            data.InitialState = new DockablePaneState();
            data.InitialState.DockPosition = DockPosition.Left;
        }

        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ExcelTable excel = this.comboBox.SelectedItem as ExcelTable;
            List<CEntityName> all_entity = App.Instance.Sqlite.SelectAllEntities(excel.Signal);
            this.dataGrid.ItemsSource = all_entity;
            
        }
    }
    
}
