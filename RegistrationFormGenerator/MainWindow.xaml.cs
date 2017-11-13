using System.Windows;
using System.Timers;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.Win32;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using RegistrationFormGenerator.Library;

namespace RegistrationFormGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonChooseExcelFile_Click(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Files (*.xls)|*.xlsx|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif" };
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Files (*.xls)|*.xlsx" };
            var result = ofd.ShowDialog();
            if (result == false)
                return;
            else
                ExcelFileLocation.Text = ofd.FileName;
        }

        private void ButtonImageFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result.ToString().Equals("Ok"))
                ImageFolderLocation.Text = dialog.FileName;
            else
                return;
        }

        private void ButtonOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result.ToString().Equals("Ok"))
                OutputFolderLocation.Text = dialog.FileName;
            else
                return;
        }

        private void ButtonGenerate_Click(object sender, RoutedEventArgs e)
        {
            List<ExcelDataRow> excelDataList;

            if (ExcelFileLocation.Text.Length< 5 || ImageFolderLocation.Text.Length < 3 || OutputFolderLocation.Text.Length < 3)
            {
                MessageBox.Show("Please choose file and folder");
                return;
            }
            ProgressPdfCreatePercenage.Value = 0;
            ProgressPdfCreatePercenage.Visibility = Visibility.Visible;

            excelDataList = ExcelReader.GeneratePdfReport(ExcelFileLocation.Text);

            foreach (ExcelDataRow data in excelDataList)
            {
                ProgressPdfCreatePercenage.Value += 1/excelDataList.Count*100;  //Progress bar increament
                ExcelPdfGenerator.GeneratePdf(data, ImageFolderLocation.Text,OutputFolderLocation.Text);      //Generate PDF Here
            }

            ResetFields();
        }

        private void ResetFields()
        {
            DeleteAllTempFiles(OutputFolderLocation.Text);  //Delete All HTML File
            ExcelFileLocation.Text = "";
            ImageFolderLocation.Text = "";
            OutputFolderLocation.Text = "";

            ProgressPdfCreatePercenage.Value = 100;
            MessageBox.Show("All Generation Done");
            ProgressPdfCreatePercenage.Visibility = Visibility.Collapsed;
        }

        private void DeleteAllTempFiles(string folderLocation)
        {
            List<FileInfo> files = GetFiles(folderLocation, ".html", ".obj");

            foreach (FileInfo file in files)
                try
                {
                    file.Attributes = FileAttributes.Normal;
                    File.Delete(file.FullName);
                }
                catch { }
        }

        private List<FileInfo> GetFiles(string path, params string[] extensions)
        {
            List<FileInfo> list = new List<FileInfo>();
            foreach (string ext in extensions)
                list.AddRange(new DirectoryInfo(path).GetFiles("*" + ext).Where(p =>
                      p.Extension.Equals(ext, StringComparison.CurrentCultureIgnoreCase))
                      .ToArray());
            return list;
        }
    }
}
