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
        bool _locker = false;
        public MainWindow()
        {
            if (Properties.Settings.Default.FirstRun == false)
            {
                Properties.Settings.Default.FirstRun = true;

                Properties.Settings.Default.FacultyName = FacultyName.ArtsandHumanities;

                Properties.Settings.Default.BengaliTextBusinessStudies = Properties.Resources.BengaliTextBusinessStudies;
                Properties.Settings.Default.EnglishTextBusinessStudies = Properties.Resources.EnglishTextBusinessStudies;

                Properties.Settings.Default.BengaliTextArtsAndHumanities = Properties.Resources.BengaliTextArtsAndHumanities;
                Properties.Settings.Default.EnglishTextArtsAndHumanities = Properties.Resources.EnglishTextArtsAndHumanities;

                Properties.Settings.Default.BengaliTextBioSciences = Properties.Resources.BengaliTextBioSciences;
                Properties.Settings.Default.EnglishTextBioSciences = Properties.Resources.EnglishTextBioSciences;

                Properties.Settings.Default.BengaliTextLaw = Properties.Resources.BengaliTextLaw;
                Properties.Settings.Default.EnglishTextLaw = Properties.Resources.EnglishTextLaw;

                Properties.Settings.Default.BengaliTextScienceAndEngineering = Properties.Resources.BengaliTextScienceAndEngineering;
                Properties.Settings.Default.EnglishTextScienceAndEngineering = Properties.Resources.EnglishTextScienceAndEngineering;

                Properties.Settings.Default.BengaliTextSocialSciences = Properties.Resources.BengaliTextSocialSciences;
                Properties.Settings.Default.EnglishTextSocialSciences = Properties.Resources.EnglishTextSocialSciences;

                Properties.Settings.Default.Save();
            }

            InitializeComponent();

            //Set Config
            switch (Properties.Settings.Default.FacultyName)
            {
                case FacultyName.ArtsandHumanities:
                    RadioArtsAndHumanities.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextArtsAndHumanities;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextArtsAndHumanities;
                    break;
                case FacultyName.BusinessStudies:
                    RadioBusinessStudies.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBusinessStudies;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBusinessStudies;
                    break;
                case FacultyName.BioScience:
                    RadioBioSciences.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBioSciences;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBioSciences;
                    break;
                case FacultyName.Law:
                    RadioLaw.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextLaw;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextLaw;
                    break;
                case FacultyName.ScienceAndEngineering:
                    RadioScienceAndEngineering.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextScienceAndEngineering;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextScienceAndEngineering;
                    break;
                case FacultyName.SocialSciences:
                    RadioSocialSciences.IsChecked = true;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextSocialSciences;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextSocialSciences;
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
            #if DEBUG
                MessageBox.Show("Done");
            #else
                ResetFields();
            #endif
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
            _locker = true;
            RadioButton button = sender as RadioButton;

            switch (button.Name)
            {
                case "RadioArtsAndHumanities":
                    RadioArtsAndHumanities.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.ArtsandHumanities;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextArtsAndHumanities;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextArtsAndHumanities;
                    break;
                case "RadioBioSciences":
                    RadioBioSciences.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.BioScience;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBioSciences;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBioSciences;
                    break;
                case "RadioBusinessStudies":
                    RadioBusinessStudies.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.BusinessStudies;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextBusinessStudies;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextBusinessStudies;
                    break;
                case "RadioLaw":
                    RadioLaw.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.Law;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextLaw;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextLaw;
                    break;
                case "RadioScienceAndEngineering":
                    RadioScienceAndEngineering.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.ScienceAndEngineering;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextScienceAndEngineering;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextScienceAndEngineering;
                    break;
                case "RadioSocialSciences":
                    RadioSocialSciences.IsChecked = true;
                    Properties.Settings.Default.FacultyName = FacultyName.SocialSciences;
                    BengaliText.Text = Properties.Settings.Default.BengaliTextSocialSciences;
                    EnglishText.Text = Properties.Settings.Default.EnglishTextSocialSciences;
                    break;
                default:
                    // ... Display button content as title.
                    this.Title = button.Content.ToString();
                    break;
            }
            Properties.Settings.Default.Save();
            _locker = false;
        }

        private void TemplateTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_locker)
                return;
            // ... Get control that raised this event.
            TextBox textBox = sender as TextBox;

            switch (Properties.Settings.Default.FacultyName)
            {
                case FacultyName.ArtsandHumanities:
                    Properties.Settings.Default.BengaliTextArtsAndHumanities = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextArtsAndHumanities = EnglishText.Text;
                    break;
                case FacultyName.BioScience:
                    Properties.Settings.Default.BengaliTextBioSciences = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextBioSciences = EnglishText.Text;
                    break;
                case FacultyName.BusinessStudies:
                    Properties.Settings.Default.BengaliTextBusinessStudies = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextBusinessStudies = EnglishText.Text;
                    break;
                case FacultyName.Law:
                    Properties.Settings.Default.BengaliTextLaw = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextLaw = EnglishText.Text;
                    break;
                case FacultyName.ScienceAndEngineering:
                    Properties.Settings.Default.BengaliTextScienceAndEngineering = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextScienceAndEngineering = EnglishText.Text;
                    break;
                case FacultyName.SocialSciences:
                    Properties.Settings.Default.BengaliTextSocialSciences = BengaliText.Text;
                    Properties.Settings.Default.EnglishTextSocialSciences = EnglishText.Text;
                    break;
                default:
                    this.Title = textBox.Name.ToString();
                    break;
            }
            Properties.Settings.Default.Save();
        }
    }
}
