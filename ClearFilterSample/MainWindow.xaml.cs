#region Copyright Syncfusion Inc. 2001 - 2015
// Copyright Syncfusion Inc. 2001 - 2015. All rights reserved.
// Use of this code is subject to the terms of our license.
// A copy of the current license can be obtained at any time by e-mailing
// licensing@syncfusion.com. Any infringement will be prosecuted under
// applicable laws. 
#endregion

using System.Windows;
using System.IO;
using System;
using System.Data;
using Syncfusion.Windows.Tools.Controls;

namespace SpreadsheetDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
         
            spreadsheetControl.Open(@"..\\..\\sheet.xlsx");
            this.spreadsheetControl.WorkbookLoaded += SpreadsheetControl_WorkbookLoaded;
        }

        private void SpreadsheetControl_WorkbookLoaded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorkbookLoadedEventArgs args)
        {
            //To clear filters in the excelsheet while loading
            if (spreadsheetControl.ActiveSheet.AutoFilters.FilterRange != null)
            {
               spreadsheetControl.ActiveSheet.AutoFilters.FilterRange = null;

            }
        }

        /// <summary>
        /// Provide support for Excel like closing operation when press the close button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibbonWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.spreadsheetControl.Commands.FileClose.Execute(null);
            if (Application.Current.ShutdownMode != ShutdownMode.OnExplicitShutdown)
                e.Cancel = true;
        }
    }
}
