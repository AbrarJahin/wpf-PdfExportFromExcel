﻿using HtmlAgilityPack;
using System.IO;
using System.Windows;
using System;
using PdfSharp;
using PdfSharp.Pdf;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace RegistrationFormGenerator.Library
{
    class ExcelPdfGenerator
    {
        internal static bool GeneratePdf(ExcelDataRow data, string outputFolderLocation)
        {
            string htmlFileLocation = GenerateHtml(data, outputFolderLocation);
            MessageBox.Show(data.Serial);
            return GeneratePdf(htmlFileLocation, @"D:\path.pdf");
        }

        private static bool GeneratePdf(string htmlFileLocation, string outputPdflocation)
        {
            string html = File.ReadAllText(htmlFileLocation);
            PdfDocument pdf = PdfGenerator.GeneratePdf(html, PageSize.A4);
            pdf.Save(outputPdflocation);
            return true;
        }

        private static string GenerateHtml(ExcelDataRow data, string outputFolderLocation)
        {
            string htmlTemplate = Properties.Resources.Html_Template;

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(htmlTemplate);

            //Update html
            htmlDocument.GetElementbyId("RegistrationNo").InnerHtml = data.RegistrationNo;
            htmlDocument.GetElementbyId("DepertmentBengali").InnerHtml = data.DepertmentBengali;
            htmlDocument.GetElementbyId("DepertmentEnglish").InnerHtml = data.DepertmentEnglish;
            //htmlDocument.GetElementbyId("Image").InnerHtml = data.DepertmentEnglish;
            htmlDocument.GetElementbyId("NameBengali").InnerHtml = data.NameBengali;
            htmlDocument.GetElementbyId("NameEnglish").InnerHtml = data.NameEnglish;
            htmlDocument.GetElementbyId("DateOfBirth").InnerHtml = data.DateOfBirth;
            htmlDocument.GetElementbyId("FatherNameBengali").InnerHtml = data.FatherNameBengali;
            htmlDocument.GetElementbyId("FatherNameEnglish").InnerHtml = data.FatherNameEnglish;
            htmlDocument.GetElementbyId("MotherNameBengali").InnerHtml = data.MotherNameBengali;
            htmlDocument.GetElementbyId("MotherNameEnglish").InnerHtml = data.MotherNameEnglish;
            htmlDocument.GetElementbyId("Session").InnerHtml = data.Session;


            string html = htmlDocument.DocumentNode.OuterHtml;

            string outputFileLocation = outputFolderLocation + "\\" + data.Serial + ".html";
            //File.WriteAllText(outputFileLocation, html);
            File.WriteAllText(@"D:\path.html", html);

            //return outputFileLocation;
            return @"D:\path.html";
        }
    }
}
