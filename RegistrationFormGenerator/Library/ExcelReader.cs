using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace RegistrationFormGenerator.Library
{
    class ExcelReader
    {
        static List<ExcelDataRow> excelDataList = new List<ExcelDataRow>();

        public static List<ExcelDataRow> GeneratePdfReport(string excelFileLocation)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(@excelFileLocation);

            ReadBySheets(xlWorkbook);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return excelDataList;
        }

        private static void ReadBySheets(Workbook xlWorkbook)
        {
            for (int sheetNo = 1; sheetNo <= xlWorkbook.Sheets.Count; sheetNo++)
            {
                _Worksheet xlWorksheet = xlWorkbook.Sheets[sheetNo];
                Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 2; i <= rowCount; i++)     //Should skip first Row, so start with 2
                {
                    ExcelDataRow row = new ExcelDataRow();
                    string cellValue;
                    for (int j = 1; j <= colCount; j++)
                    {
                        try
                        {
                            cellValue = xlWorksheet.UsedRange.Cells[i, j].Value.ToString();
                        }
                        catch (Exception ex)
                        {
                            cellValue = "";
                        }

                        switch (j)
                        {
                            case 1:
                                row.Serial = sheetNo + cellValue;
                                break;
                            case 2:
                                row.NameBengali = cellValue;
                                break;
                            case 3:
                                row.NameEnglish = cellValue;
                                break;
                            case 4:
                                row.DateOfBirth = cellValue;
                                break;
                            case 5:
                                row.RegistrationNo = cellValue;
                                break;
                            case 6:
                                row.RollNo = cellValue;
                                break;
                            case 7:
                                row.SessionBengali = cellValue;
                                break;
                            case 8:
                                row.SessionEnglish = cellValue;
                                break;
                            case 9:
                                row.FatherNameBengali = cellValue;
                                break;
                            case 10:
                                row.FatherNameEnglish = cellValue;
                                break;
                            case 11:
                                row.MotherNameBengali = cellValue;
                                break;
                            case 12:
                                row.MotherNameEnglish = cellValue;
                                break;
                            case 13:
                                row.MobileNo = cellValue;
                                break;
                            case 14:
                                row.PresentAddress = cellValue;
                                break;
                            case 15:
                                row.PermanentAddress = cellValue;
                                break;
                            case 16:
                                int parsedInt = 0;
                                if (int.TryParse(cellValue, out parsedInt))
                                {
                                    if(parsedInt<1 || parsedInt>6)
                                        parsedInt = 0;
                                }
                                row.TemplateAutoSelectCode = parsedInt;
                                break;
                            case 17:
                                row.FacultyBengali = cellValue;
                                break;
                            case 18:
                                row.FacultyEnglish = cellValue;
                                break;
                            case 19:
                                row.DepertmentBengali = cellValue;
                                break;
                            case 20:
                                row.DepertmentEnglish = cellValue;
                                break;
                            case 21:
                                row.DegreeNameBengali = cellValue;
                                break;
                            case 22:
                                row.DegreeNameEnglish = cellValue;
                                break;
                            case 23:
                                row.AdmissionCancelled = cellValue;
                                break;
                            case 24:
                                row.Comment = cellValue;
                                break;
                            default:
                                break;
                        }
                    }
                    excelDataList.Add(row);
                }

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
            }
        }
    }
}
