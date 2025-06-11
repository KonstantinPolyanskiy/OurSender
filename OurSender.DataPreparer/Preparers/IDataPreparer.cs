using OurSender.DataPreparer.Models;

namespace OurSender.DataPreparer.Preparers;

/// <summary>
/// Принимает данные и структурирует их в виде <see cref="PreparedData"/>
/// </summary>
public interface IDataPreparer
{
    /// <summary>
    /// Структурирует данные из потока
    /// </summary>
    Task<PreparedData> PrepareAsync(Stream stream, CancellationToken ct = default);
}