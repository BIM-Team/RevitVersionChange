﻿using Autodesk.Revit.UI;
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
        public ImageControl()
        {
            InitializeComponent();
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

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
    }
}
