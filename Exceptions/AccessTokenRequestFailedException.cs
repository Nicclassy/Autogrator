namespace Autogrator.Exceptions;

[Serializable]
public sealed class AccessTokenRequestFailedException : Exception {
    public AccessTokenRequestFailedException() { }

    public AccessTokenRequestFailedException(string message) : base(message) { }

    public AccessTokenRequestFailedException(string message, Exception innerException) : base(message, innerException) { }
}