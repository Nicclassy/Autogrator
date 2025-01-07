using System;
using System.Diagnostics;

using Serilog;

using Autogrator.Utilities;

public static class OutlookUtilities {
    private static readonly string CreateProfileArgs = "/profile";

    public static void CreateProfile(string profileName) {
        // Only works on classic Outlook
        ProcessStartInfo info = new() {
            FileName = Directories.OutlookExecutable,
            Arguments = CreateProfileArgs + $" {profileName}"
        };
        using Process process = Process.Start(info)!;
        process.WaitForExit();
        if (process.ExitCode != 0)
            Log.Warning(
                "Profile creation process did not exit with code 0, " +
                $"instead exited with code {process.ExitCode}"
            );
    }
}
