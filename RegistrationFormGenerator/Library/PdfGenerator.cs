using HtmlAgilityPack;
using System.Windows;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.IO;
using System;
using iTextSharp.text.html;
using System.Collections.Generic;

namespace RegistrationFormGenerator.Library
{
    class ExcelPdfGenerator
    {
        private static string currentRegistrationNo;
        internal static bool GenerateHtmlPdf(ExcelDataRow data, string imageFolderLocation, string outputFolderLocation)
        {
            string htmlString = GenerateHtml(data,imageFolderLocation);
            currentRegistrationNo = data.RegistrationNo;
            return GenerateHtmlPdf(htmlString, @outputFolderLocation + "\\" + data.Serial+".pdf");
        }

        private static bool GenerateHtmlPdf(string htmlString, string outputPdflocation)
        {
            bool ifCreatedSuccessfully = true;
            Document doc = new Document(PageSize.A3);
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(outputPdflocation, FileMode.Create));
            doc.Open();

            try
            {
                //Path to our font
                string arialuniTff = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                    "ARIALUNI.TTF");
                //Register the font with iTextSharp
                FontFactory.Register(arialuniTff);

                //Create a new stylesheet
                StyleSheet ST = new StyleSheet();
                //Set the default body font to our registered font's internal name
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.FACE, "Arial Unicode MS");
                //Set the default encoding to support Unicode characters
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.ENCODING, BaseFont.IDENTITY_H);

                //Parse our HTML using the stylesheet created above
                List<IElement> list = HTMLWorker.ParseToList(new StringReader(htmlString), ST);

                //Loop through each element, don't bother wrapping in P tags
                foreach (var element in list)
                {
                    doc.Add(element);
                }
                doc.Close();
                wri.Close();
            }
            catch (Exception ex)
            {
                ifCreatedSuccessfully = false;
                Console.WriteLine(ex);
                MessageBox.Show("PDF Generation Failed for RegNo - "+ currentRegistrationNo);
                //throw;
            }
            finally
            {
                //should use this to show unicode - http://www.codescratcher.com/asp-net/display-unicode-characters-in-converting-html-to-pdf/
            }
            return ifCreatedSuccessfully;
        }

        private static string GenerateHtml(ExcelDataRow data, string imageFolderLocation)
        {
            string htmlTemplate = Properties.Resources.Html_Template;

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(htmlTemplate);

            //Update html
            htmlDocument.GetElementbyId("RegistrationNo").InnerHtml = data.RegistrationNo;
            htmlDocument.GetElementbyId("DepertmentBengali").InnerHtml = data.DepertmentBengali;
            htmlDocument.GetElementbyId("DepertmentEnglish").InnerHtml = data.DepertmentEnglish;
            htmlDocument.GetElementbyId("Image").SetAttributeValue("src", imageFolderLocation +"\\"+ data.RegistrationNo+".jpg");
            htmlDocument.GetElementbyId("NameBengali").InnerHtml = data.NameBengali;
            htmlDocument.GetElementbyId("NameEnglish").InnerHtml = data.NameEnglish;
            htmlDocument.GetElementbyId("DateOfBirth").InnerHtml = data.DateOfBirth;
            htmlDocument.GetElementbyId("FatherNameBengali").InnerHtml = data.FatherNameBengali;
            htmlDocument.GetElementbyId("FatherNameEnglish").InnerHtml = data.FatherNameEnglish;
            htmlDocument.GetElementbyId("MotherNameBengali").InnerHtml = data.MotherNameBengali;
            htmlDocument.GetElementbyId("MotherNameEnglish").InnerHtml = data.MotherNameEnglish;
            htmlDocument.GetElementbyId("Session").InnerHtml = data.Session;

            return htmlDocument.DocumentNode.OuterHtml;
        }
    }
}
