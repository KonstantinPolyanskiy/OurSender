using ClosedXML.Excel;

namespace OurSender.Tests.DataPreparerTests.ExcelDataPreparer;

internal static class ExcelHelper
{
    public static byte[] ExcelAsBytes(
        string[] headers,
        params object?[][] rows)
    {
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var worksheet = wb.Worksheets.Add("Sheet1");
            
            for (var c = 0; c < headers.Length; c++)
                worksheet.Cell(1, c + 1).Value = headers[c];

            for (var r = 0; r < rows.Length; r++)
            {
                var row = rows[r];
                for (var c = 0; c < row.Length; c++)
                {
                    var cell = worksheet.Cell(r + 2, c + 1);
                    dynamic? v = row[c];
                    if (v != null)
                        cell.SetValue(v);           
                }
            }

            wb.SaveAs(ms);
        }
        
        return ms.ToArray();
    }
}