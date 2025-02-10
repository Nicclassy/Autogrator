using Serilog;

using Autogrator.Utilities;

namespace Autogrator.Notifications;

public enum TimeNotificationInterval {
    Minute,
    Hour,
    Day
}

public sealed class TimeNotifier {
    public readonly TimeNotificationInterval interval;
    public DateTime notificationTime;

    public string Message { get; init; } = "This is a scheduled time notification";
    public IAnsiSequence Colour { get; init; } = AnsiColours.BrightGreen;

    public TimeNotifier(TimeNotificationInterval interval) {
        this.interval = interval;
        notificationTime = InitialNotificationTime();
    }

    public void NotifyIfTime() {
        if (DateTime.Now >= notificationTime) {
            Log.Information("{Colour}{Message}{Reset}", Colour, Message, AnsiColours.Reset);
            notificationTime = NextNotificationTime(notificationTime);
        }
    }

    private DateTime InitialNotificationTime() {
        DateTime now = DateTime.Now;
        DateTime DateTimeWithProperties(int? day = null, int? hour = null, int? minute = null, int? second = null) =>
            new DateTime(
                now.Year, now.Month, 
                day ?? now.Day, 
                hour ?? now.Hour, 
                minute ?? now.Minute, 
                second ?? now.Second
            );

        DateTime previous = interval switch {
            TimeNotificationInterval.Minute => DateTimeWithProperties(second: 0),
            TimeNotificationInterval.Hour => DateTimeWithProperties(minute: 0, second: 0),
            TimeNotificationInterval.Day => DateTimeWithProperties(hour: 0, minute: 0, second: 0),
            _ => throw new ArgumentException($"The interval {interval.ToString()} is not a supported interval")
        };
        return NextNotificationTime(previous);
    }

    public DateTime NextNotificationTime(DateTime previous) => interval switch {
        TimeNotificationInterval.Minute => previous + TimeSpan.FromMinutes(1),
        TimeNotificationInterval.Hour => previous + TimeSpan.FromHours(1),
        TimeNotificationInterval.Day => previous + TimeSpan.FromDays(1),
        _ => throw new ArgumentException($"The interval {interval.ToString()} is not a supported interval")
    };
}