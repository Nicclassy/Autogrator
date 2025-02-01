namespace Autogrator.Exceptions;

[Serializable]
public sealed class AppDataNotFoundException : Exception {
    public AppDataNotFoundException() {}

    public AppDataNotFoundException(string message) : base(message) {}

    public AppDataNotFoundException(string message, Exception innerException) : base(message, innerException) {}
}