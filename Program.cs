using Serilog;

using Autogrator.OutlookAutomation;

namespace Autogrator;

public static class Program {
    static Program() {
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .CreateLogger();
        Log.Information("Begin logging...");
    }
    public static void Main() {
       var receiver = OutlookEmailReceiver.Create();
       if (receiver.TryReceiveEmail(out var email)) {
            Log.Information($"Received information: {email}");
       } else {
            Log.Information("Did not receive anything.");
       }
       Log.CloseAndFlush();
    }
}