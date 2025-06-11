using FluentAssertions;

namespace OurSender.Tests.DataPreparerTests.ExcelDataPreparer.ExcelDataPreparerTests;

public class ExcelPreparerTests
{
    [Fact]
    public async Task PrepareAsync_WhenSingleDataRow_ReturnCorrectDictionary_Ok()
    {
        var excelData = ExcelHelper.ExcelAsBytes(
            headers: ["name"],
            rows: new object?[][]
            {
                ["Константин"],
            }
        );
        
        using var excelStream = new MemoryStream(excelData);

        var preparer = new DataPreparer.Preparers.Implementations.ExcelDataPreparer();

        var result = await preparer.PrepareAsync(excelStream);
        var list = result.Rows.ToList();
        
        list.Should().HaveCount(1);
        
        list[0].Should()
            .ContainKey("name")
            .WhoseValue.ToString().Should().Be("Константин");
    }

    [Fact]
    public async Task PrepareAsync_WhenMultipleDataRowsAndColumns_ReturnCorrectDictionary_Ok()
    {
        var excelData = ExcelHelper.ExcelAsBytes(
            headers: ["name", "phone", "debt"],
            rows: new object?[][]
            {
                ["Константин", "+7100", 100],
                ["Анатолий", "+7200", 200],
                ["Иван", "+7300", 300]
            });
        
        using var excelStream = new MemoryStream(excelData);
        
        var preparer = new DataPreparer.Preparers.Implementations.ExcelDataPreparer();
        
        var result = await preparer.PrepareAsync(excelStream);
        var list = result.Rows.ToList();
        
        list.Should().HaveCount(3);

        list[0].Should().BeEquivalentTo(new Dictionary<string, object>
        {
            ["name"]  = "Константин",
            ["phone"] = "+7100",
            ["debt"]  = 100,
        });

        list[1].Should().BeEquivalentTo(new Dictionary<string, object>
        {
            ["name"]  = "Анатолий",
            ["phone"] = "+7200",
            ["debt"]  = 200
        });

        list[2].Should().BeEquivalentTo(new Dictionary<string, object>
        {
            ["name"]  = "Иван",
            ["phone"] = "+7300",
            ["debt"]  = 300
        });
        
    }

    [Fact]
    public async Task PrepareAsync_WhenMultipleDataRowsAndColumnsAndNull_ReturnCorrectDictionary_Ok()
    {
        var excelData = ExcelHelper.ExcelAsBytes(
            headers: ["name", "phone", "debt"],
            rows: new object?[][]
            {
                [null, "+7100", 100],
                ["Анатолий", null, 200],
                ["Иван", "+7300", null]
            });
        
        using var excelStream = new MemoryStream(excelData);
        
        var preparer = new DataPreparer.Preparers.Implementations.ExcelDataPreparer();
        
        var result = await preparer.PrepareAsync(excelStream);
        var list = result.Rows.ToList();
        
        list.Should().HaveCount(3);
        
        list[0].Should().BeEquivalentTo(new Dictionary<string, object?>
            {
                ["name"]  = string.Empty,
                ["phone"] = "+7100",
                ["debt"]  = 100
            }, options => options.WithStrictOrdering());

        list[1].Should().BeEquivalentTo(new Dictionary<string, object?>
        {
            ["name"]  = "Анатолий",
            ["phone"] = string.Empty,
            ["debt"]  = 200
        }, options => options.WithStrictOrdering());

        list[2].Should().BeEquivalentTo(new Dictionary<string, object?>
        {
            ["name"]  = "Иван",
            ["phone"] = "+7300",
            ["debt"]  = string.Empty
        }, options => options.WithStrictOrdering());
    }

    [Fact]
    public async Task PrepareAsync_WhenManyDataRows_ReturnCorrectDictionary_Ok()
    {
        var excelData = ExcelHelper.ExcelAsBytes(
            headers: ["name"],
            rows: new object?[][]
            {
                ["Константин"],
                ["Анатолий"],
                ["Иван"]
            }
        );
        
        using var excelStream = new MemoryStream(excelData);

        var preparer = new DataPreparer.Preparers.Implementations.ExcelDataPreparer();
        
        var result = await preparer.PrepareAsync(excelStream);
        var list = result.Rows.ToList();
        
        list.Should().HaveCount(3);

        var names = list
            .Select(dict => dict["name"].ToString())  
            .ToList();

        names.Should().Equal("Константин", "Анатолий", "Иван");
    }
    
    [Fact]
    public async Task PrepareAsync_WhenEmptyWorksheet_ThrowsNullReferenceException()
    {
        var excelBytes = ExcelHelper.ExcelAsBytes([]);

        using var excelStream = new MemoryStream(excelBytes);
        
        var preparer = new DataPreparer.Preparers.Implementations.ExcelDataPreparer();
        
        Func<Task> act = () => preparer.PrepareAsync(excelStream);

        await act.Should().ThrowAsync<NullReferenceException>();
    }
    
}