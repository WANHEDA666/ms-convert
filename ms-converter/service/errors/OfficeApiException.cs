namespace ms_converter.service.errors;

public sealed class OfficeApiException(string message) : Exception(message);