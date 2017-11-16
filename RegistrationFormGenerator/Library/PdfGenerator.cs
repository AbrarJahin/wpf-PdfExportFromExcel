using HtmlAgilityPack;
using System.IO;

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
            //Save in HTML
            File.WriteAllText(outputPdflocation+".html",htmlString );
            /*
            //Should add embaded Image - https://stackoverflow.com/a/19398426/2193439
            //Add Bengla Text - https://www.codeproject.com/Questions/1150398/How-do-I-write-bengali-in-pdfptable-using-iTextsha
            //PDFSharp - https://stackoverflow.com/a/31109987/2193439
            //Adding Unicode - https://stackoverflow.com/a/31606661/2193439
            Document pdfDoc = new Document(PageSize.A3);
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, new FileStream(outputPdflocation, FileMode.Create));
            pdfDoc.Open();

            try
            {
                //Path to our font
                string solaimanLipiTff = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                    "SolaimanLipi.ttf");
                //Register the font with iTextSharp
                FontFactory.Register(solaimanLipiTff);

                //Register SolaimanLipi font
                //FontFactory.Register(Encoding.Unicode.GetString(Properties.Resources.SolaimanLipi));

                //Create a new stylesheet
                StyleSheet ST = new StyleSheet();
                //Set the default body font to our registered font's internal name
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.FACE, "SolaimanLipi");
                //Set the default encoding to support Unicode characters
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.ENCODING, BaseFont.IDENTITY_H);

                //Parse our HTML using the stylesheet created above
                List<IElement> list = HTMLWorker.ParseToList(new StringReader(htmlString), ST);

                //Loop through each element, don't bother wrapping in P tags
                foreach (var element in list)
                {
                    pdfDoc.Add(element);
                }
                pdfDoc.Close();
                pdfWriter.Close();
            }
            catch (Exception ex)
            {
                ifCreatedSuccessfully = false;
                if (Debugger.IsAttached == true)
                {
                    MessageBox.Show(ex.StackTrace);
                }
                else
                {
                    MessageBox.Show("Image not found for RegNo - " + currentRegistrationNo);
                }
                //throw;
            }
            finally
            {
                //should use this to show unicode - http://www.codescratcher.com/asp-net/display-unicode-characters-in-converting-html-to-pdf/
            }
            */
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
