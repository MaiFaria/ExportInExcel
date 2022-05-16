using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace ExportInExcel
{
    public static class ExportInExcel
    {
        public static string SimpleExportExcel<T>(List<T> list, string sheetName)
        {
            using (IXLWorkbook wb = new XLWorkbook())
            {
                var workbook = wb.AddWorksheet(sheetName).FirstCell().InsertTable<T>(list, false);

                workbook.Rows(1, 1).Style.Font.Bold = true;
                workbook.Rows(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);

                    return Convert.ToBase64String(stream.ToArray());
                }
            }
        }

        public static string ExportExcelWithCustomTitle<T>(List<T> list, List<string> titleList, string sheetName)
        {
            using (IXLWorkbook wb = new XLWorkbook())
            {
                var workbook = wb.AddWorksheet(sheetName).FirstCell().InsertTable<T>(list, false);
                var currentRow = 1;
                var Column = 1;

                workbook.Rows(1, 1).Style.Font.Bold = true;
                workbook.Rows(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                foreach (var row in titleList)
                {
                    workbook.Cell(currentRow, Column++).Value = row;
                }

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);

                    return Convert.ToBase64String(stream.ToArray());
                }
            }
        }
    }
}
