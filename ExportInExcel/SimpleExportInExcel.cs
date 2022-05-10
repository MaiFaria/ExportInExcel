using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace ExportInExcel
{
    public static class SimpleExportInExcel
    {
        public static string ExportInExcel<T>(List<T> list, string sheetName)
        {
            using (IXLWorkbook wb = new XLWorkbook())
            {
                var workbook = wb.AddWorksheet(sheetName).FirstCell().InsertTable<T>(list, false);
                workbook.Rows(1, 1).Style.Font.Bold = true;

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);

                    return Convert.ToBase64String(stream.ToArray());
                }
            }
        }
    }
}
