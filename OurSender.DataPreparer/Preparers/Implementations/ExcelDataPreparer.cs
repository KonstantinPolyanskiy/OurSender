using System.Collections.Immutable;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OurSender.DataPreparer.Models;
using OurSender.DataPreparer.Models.Exceptions;

namespace OurSender.DataPreparer.Preparers.Implementations;

public class ExcelDataPreparer : IDataPreparer
{
    public Task<PreparedData> PrepareAsync(Stream stream, CancellationToken ct = default)
    {
        using var wb = new XLWorkbook(stream);
        
        var worksheet = wb.Worksheets.First();
        
        var headers = worksheet.Row(1)
            .Cells()
            .Select(cell => cell.Value.ToString().Trim())
            .ToImmutableArray();
        
        var range = worksheet.RangeUsed();
        if (range is null)
            throw new NullReferenceException("worksheet is empty");
        
        var preparedRows = new List<IDictionary<string, object>>();

        foreach (var dataInRow in range.RowsUsed().Skip(1))
        {
            ct.ThrowIfCancellationRequested();
            
            var dict = new Dictionary<string, object>(headers.Length);

            for (var i = 0; i < headers.Length; i++)
            {
                var cell = dataInRow.Cell(i + 1);

                object value = cell.DataType switch
                {
                    XLDataType.Blank => string.Empty,
                    XLDataType.Boolean => cell.GetBoolean(),
                    XLDataType.Number => cell.GetDouble(),
                    XLDataType.Text => cell.GetString(),
                    XLDataType.Error => cell.GetString(),
                    XLDataType.DateTime => cell.GetDateTime(),
                    XLDataType.TimeSpan => cell.GetTimeSpan(),
                    _ => throw new ArgumentOutOfRangeException()
                };
                
                dict.Add(headers[i], value);
            }
            
            preparedRows.Add(dict);
        }

        return Task.FromResult(new PreparedData { Rows = preparedRows });
    }
}