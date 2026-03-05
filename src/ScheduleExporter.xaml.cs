using Microsoft.Win32;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace ScheduleExporter {
    public partial class ExportWindow : Window {
        public string SelectedFilePath { get; private set; }
        public bool WriteAsString { get; private set; }
        public bool ExcludeEmptyRows { get; private set; }
        public List<string> SelectedSchedules { get; private set; }

        private ObservableCollection<ScheduleItem> scheduleItems;

        public ExportWindow(List<string> scheduleNames) {
            InitializeComponent();
            PopulateSchedules(scheduleNames);
        }

        private void PopulateSchedules(List<string> scheduleNames) {
            scheduleItems = new ObservableCollection<ScheduleItem>();

            foreach (var schedule in scheduleNames) {
                scheduleItems.Add(new ScheduleItem { Name = schedule, IsChecked = false });
            }

            listBoxSchedules.ItemsSource = scheduleItems;
        }

        private void ButtonSelectFile_Click(object sender, RoutedEventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true) {
                textBoxFilePath.Text = openFileDialog.FileName;
            }
        }

        private void ButtonNewFile_Click(object sender, RoutedEventArgs e) {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Save an Excel File",
                FileName = "SchedulesExport"
            };

            if (saveFileDialog.ShowDialog() == true) {
                try {
                    using (var workbook = new XLWorkbook()) {
                        workbook.Worksheets.Add("Sheet1");
                        workbook.SaveAs(saveFileDialog.FileName);
                    }
                    textBoxFilePath.Text = saveFileDialog.FileName;
                }
                catch (Exception ex) {
                    MessageBox.Show("Error: " + ex.Message, "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ButtonWrite_Click(object sender, RoutedEventArgs e) {
            SelectedFilePath = textBoxFilePath.Text;
            WriteAsString = checkBoxWriteAsString.IsChecked == true;
            ExcludeEmptyRows = checkBoxExcludeEmpty.IsChecked == true;

            SelectedSchedules = scheduleItems
                .Where(item => item.IsChecked)
                .Select(item => item.Name)
                .ToList();

            if (string.IsNullOrWhiteSpace(SelectedFilePath) || SelectedSchedules.Count == 0) {
                MessageBox.Show(
                    "Please create or select a file and at least one schedule.",
                    "Input Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            if (Helper.IsFileOpen(textBoxFilePath.Text)) {
                MessageBox.Show(
                    "The selected file is currently open. Please close it before proceeding.",
                    "File In Use",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void CheckBoxSelectAll_Changed(object sender, RoutedEventArgs e) {
            bool isChecked = checkBoxSelectAll.IsChecked == true;

            foreach (var item in scheduleItems) {
                item.IsChecked = isChecked;
            }
        }
    }

    public class ScheduleItem : INotifyPropertyChanged {
        private string name;
        private bool isChecked;

        public string Name {
            get => name;
            set {
                if (name != value) {
                    name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public bool IsChecked {
            get => isChecked;
            set {
                if (isChecked != value) {
                    isChecked = value;
                    OnPropertyChanged(nameof(IsChecked));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}