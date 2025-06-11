namespace OurSender.DataPreparer.Models.Exceptions;

/// <summary>
/// Исключение для ситуации, когда в Excel файле не найдена строка с непустым значением
/// </summary>
public sealed class NonEmptyRowNotFoundException(string message) : Exception(message);