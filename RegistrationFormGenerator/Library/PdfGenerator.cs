using System.Windows;

namespace RegistrationFormGenerator.Library
{
    class PdfGenerator
    {
        internal static void GeneratePdf(ExcelDataRow data)
        {
            string html = Properties.Resources.Html_Template;
            MessageBox.Show(html);
        }
    }
}
