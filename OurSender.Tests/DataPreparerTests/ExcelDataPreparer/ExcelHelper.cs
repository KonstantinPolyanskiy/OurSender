using ClosedXML.Excel;

namespace OurSender.Tests.DataPreparer.ExcelDataPreparer;

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
            
            
        }
    }
}