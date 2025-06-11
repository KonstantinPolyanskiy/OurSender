namespace OurSender.DataPreparer.Models;

/// <summary>
/// Подготовленные данные из источника данных
/// </summary>
public class PreparedData
{
    /// <summary>
    /// Ключ (string) - имя поля
    /// Значение (object) - содержимое поля
    /// </summary>
    public IEnumerable<IDictionary<string, object>> Rows { get; set; } = new List<IDictionary<string, object>>();
}