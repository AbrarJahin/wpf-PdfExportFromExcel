using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.Win32;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using RegistrationFormGenerator.Library;
using static RegistrationFormGenerator.Enums;
using System.Windows.Controls;

namespace RegistrationFormGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            if (Properties.Settings.Default.FirstRun == false)
            {
                Properties.Settings.Default.FirstRun = true;

                Properties.Settings.Default.FacultyName = FacultyName.AccountingAndInformation;

                Properties.Settings.Default.BengaliTextAccountingAndInformation = Properties.Resources.BengaliTextAccountingAndInformation;
                Properties.Settings.Default.EnglishTextAccountingAndInformation = Properties.Resources.EnglishTextAccountingAndInformation;

                Properties.Settings.Default.BengaliTextBangla = Properties.Resources.BengaliTextBangla;
                Properties.Settings.Default.EnglishTextBangla = Properties.Resources.EnglishTextBangla;

                Properties.Settings.Default.BengaliTextBotany = Properties.Resources.BengaliTextBotany;
                Properties.Settings.Default.EnglishTextBotany = Properties.Resources.EnglishTextBotany;

                Properties.Settings.Default.BengaliTextLaw = Properties.Resources.BengaliTextLaw;
                Properties.Settings.Default.EnglishTextLaw = Properties.Resources.EnglishTextLaw;

                Properties.Settings.Default.BengaliTextMathematics = Properties.Resources.BengaliTextMathematics;
                Properties.Settings.Default.EnglishTextMathematics = Properties.Resources.EnglishTextMathematics;

                Properties.Settings.Default.BengaliTextSociology = Properties.Resources.BengaliTextSociology;
                Properties.Settings.Default.EnglishTextSociology = Properties.Resources.EnglishTextSociology;

                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.FacultyName = FacultyName.AccountingAndInformation;
            }

            InitializeComponent();

            //Set Config
            switch (Properties.Settings.Default.FacultyName)
            {
                case FacultyName.AccountingAndInformation:
                    RadioAccountingAndInformation.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextAccountingAndInformation;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextAccountingAndInformation;
                    break;
                case FacultyName.Bangla:
                    RadioBangla.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBangla;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBangla;
                    break;
                case FacultyName.Botany:
                    RadioBotany.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBotany;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBotany;
                    break;
                case FacultyName.Law:
                    RadioLaw.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextLaw;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextLaw;
                    break;
                case FacultyName.Mathematics:
                    RadioMathematics.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextMathematics;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextMathematics;
                    break;
                case FacultyName.Sociology:
                    RadioSociology.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextSociology;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextSociology;
                    break;
            }
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
                ProgressPdfCreatePercenage.Value += (int)(100/excelDataList.Count);  //Progress bar increament
                ExcelPdfGenerator.GenerateHtmlPdf(data, ImageFolderLocation.Text,OutputFolderLocation.Text);      //Generate PDF Here
            }

            ResetFields();
        }

        private void ResetFields()
        {
            DeleteAllTempFiles(OutputFolderLocation.Text);  //Delete All HTML File
            MessageBox.Show("All Generation Done");
            Process.Start(@OutputFolderLocation.Text);
            //Hide all
            ExcelFileLocation.Text = "";
            ImageFolderLocation.Text = "";
            OutputFolderLocation.Text = "";
            ProgressPdfCreatePercenage.Visibility = Visibility.Collapsed;
        }

        private void DeleteAllTempFiles(string folderLocation)
        {
            List<FileInfo> files = GetFiles(folderLocation, ".xml", ".obj");

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

        private void FacultyChanged(object sender, RoutedEventArgs e)
        {
            RadioButton button = sender as RadioButton;

            switch (button.Name)
            {
                case "RadioAccountingAndInformation":
                    Properties.Settings.Default.FacultyName = FacultyName.AccountingAndInformation;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextAccountingAndInformation;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextAccountingAndInformation;
                    break;
                case "RadioBangla":
                    RadioBangla.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBangla;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBangla;
                    break;
                case "RadioBotany":
                    RadioBotany.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBotany;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBotany;
                    break;
                case "RadioLaw":
                    RadioLaw.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextLaw;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextLaw;
                    break;
                case "RadioMathematics":
                    RadioMathematics.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextMathematics;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextMathematics;
                    break;
                case "RadioSociology":
                    RadioSociology.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextSociology;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextSociology;
                    break;
                default:
                    // ... Display button content as title.
                    this.Title = button.Content.ToString();
                    break;
            }
        }
    }
}
