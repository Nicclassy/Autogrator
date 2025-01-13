namespace Autogrator.Exceptions;

[Serializable]
public sealed class RequestUnsuccessfulException : Exception {
    public RequestUnsuccessfulException() { }

    public RequestUnsuccessfulException(string message) : base(message) { }

    public RequestUnsuccessfulException(string message, Exception innerException) : base(message, innerException) { }
}