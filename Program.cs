using Serilog;

using Autogrator.OutlookAutomation;

namespace Autogrator;

public static class Program {
    static Program() {
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .CreateLogger();
        Log.Information("Begin logging");
    }

    public static void EmailListener() {
        var receiver = OutlookEmailReceiver.Create();
        Log.Information("Waiting for emails...");
        while (true) {
            if (receiver.TryReceiveEmail(out var email)) {
                Log.Information($"Received email: {email.Subject}");
            }
        }
    }

    public static void Main() {}
}