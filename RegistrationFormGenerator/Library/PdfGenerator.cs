using HtmlAgilityPack;
using System.Windows;

namespace RegistrationFormGenerator.Library
{
    class PdfGenerator
    {
        internal static void GeneratePdf(ExcelDataRow data)
        {
            string htmlTemplate = Properties.Resources.Html_Template;

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(htmlTemplate);

            //Update html
            HtmlNode demoNode = htmlDocument.GetElementbyId("NameBengali");
            demoNode.InnerHtml = "Try Me";
            string html = htmlDocument.DocumentNode.OuterHtml;
            MessageBox.Show(html);
        }
    }
}
