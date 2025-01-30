using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using Serilog.Events;

using Autogrator.Utilities;
using Autogrator.Notifications;

namespace Autogrator;

public partial class Autogrator {
    private static partial void SetDefaultLoggingConfiguration() {
        Dictionary<ConsoleThemeStyle, string> styles = new() {
            [ConsoleThemeStyle.SecondaryText] = new AnsiHexColour("#ce3189"),
            [ConsoleThemeStyle.TertiaryText] = new AnsiHexColour("#9c6380"),
            [ConsoleThemeStyle.Name] = new AnsiHexColour("#d70922"),
            [ConsoleThemeStyle.String] = new AnsiHexColour("#50a5af"),
            [ConsoleThemeStyle.Scalar] = new AnsiHexColour("#a05f9f"),
            [ConsoleThemeStyle.Invalid] = new AnsiHexColour("#f40b53"),
            [ConsoleThemeStyle.Null] = new AnsiHexColour("#ecf74c"),
            [ConsoleThemeStyle.Number] = new AnsiHexColour("#c63987"),
            [ConsoleThemeStyle.Boolean] = new AnsiHexColour("#dd7622"),
            [ConsoleThemeStyle.Scalar] = new AnsiHexColour("#7d5ea1"),
            [ConsoleThemeStyle.LevelInformation] = new AnsiHexColour("#5060af"),
            [ConsoleThemeStyle.LevelFatal] = AnsiColours.Red.Combine(AnsiColours.BgBrightWhite),
            [ConsoleThemeStyle.LevelDebug] = new AnsiHexColour("#ad6152"),
            [ConsoleThemeStyle.LevelWarning] = new AnsiHexColour("#ead215"),
        };
        AnsiConsoleTheme theme = new(styles);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console(theme: theme)
            .WriteTo.File(
                new StylelessTextFormatter(),
                EmailExceptionNotifier.LogFileName,
                rollingInterval: RollingInterval.Day,
                restrictedToMinimumLevel: LogEventLevel.Debug
            )
            .CreateLogger();
        Log.Information("Autogrator logging started");
    }
}