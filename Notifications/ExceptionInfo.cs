using System.Globalization;

namespace Autogrator.Notifications;

public sealed record ExceptionInfo(string ExceptionType, DateTime DateTime) {
    private const string DefaultTimeStampFormat = "f";

    public static ExceptionInfo Create(Exception ex, DateTime dateTime) => new(ex.GetType().Name, dateTime);

    public string TimeStamp(string format = DefaultTimeStampFormat) =>
        DateTime.ToString(format, CultureInfo.CurrentCulture);
}
