namespace Autogrator.Exceptions;

[Serializable]
public sealed class AccessTokenRetrievalFailedException : Exception {
    public AccessTokenRetrievalFailedException() { }

    public AccessTokenRetrievalFailedException(string message) : base(message) { }

    public AccessTokenRetrievalFailedException(string message, Exception innerException) : base(message, innerException) { }
}