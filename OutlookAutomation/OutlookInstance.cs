using System.Runtime.InteropServices;

using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public static partial class OutlookInstance {
    private const bool UseAltLogin = true;
    private const bool AutomateProfileCreation = true;

    public static Outlook.Application Application { get; } = new();
    public static Outlook.NameSpace NameSpace { get; } = Application.GetNamespace("MAPI");
    public static bool IsAuthenticated { get; private set; } = false;

    public static Outlook.MAPIFolder Inbox =>
        NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    private static string Email => 
        UseAltLogin? Credentials.Outlook.AltEmail : Credentials.Outlook.Email;
    private static string Password => 
        UseAltLogin ? Credentials.Outlook.AltPassword : Credentials.Outlook.Password;

    static OutlookInstance() => Login();

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetForegroundWindow(IntPtr hWnd);

#nullable disable
    [LibraryImport("user32.dll", EntryPoint = "FindWindowA", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial IntPtr FindWindow(string lpClassName, string lpWindowName);
#nullable enable

    public static void Login() {
        if (IsAuthenticated)
            return;

        bool retry = false;
        Log.Information("Logging in with email {Email}", Email);
        try {
            LoginWithOptions(showDialog: false, newSession: true);
        } catch (COMException) {
            Log.Warning("Initial login attempt failed. Retrying with dialog...");
            retry = true;
        }
        
        if (retry) {
            try {
                // Try again, but this time show dialog.
                // The error previously thrown may be a consequence
                // of no profile existing. Hence, showing the dialog box
                // enables the user to create a profile and thus avoid the error
                LoginWithOptions(showDialog: true, newSession: true);
            } catch (COMException ex) {
                Log.Error(ex, "Login failed");
                throw;
            }
        }

        IsAuthenticated = true;
        Log.Information("Successfully logged in!");
    }

    private static void LoginWithOptions(bool showDialog, bool newSession) {
        // Closes, so cannot use static here
        void loginAction() =>
            NameSpace.Logon(Email, Password, ShowDialog: showDialog, NewSession: newSession);

        if (!showDialog || !AutomateProfileCreation) {
            loginAction();
            return;
        }

        Thread thread = new(new ThreadStart(loginAction));
        thread.Start();

        nint window = FindWindow(null, "Choose Profile");
        if (window != nint.Zero) {
            Log.Warning("The window for profile creation was not found.");
            return;
        }

        if (!SetForegroundWindow(window)) {
            Log.Warning("The foreground window was not set.");
            return;
        }

        Console.WriteLine("Press any key to send ENTER");
        Console.ReadKey();
        Log.Information("Sending ENTER to dialog box");
        SendKeys.SendWait("{ENTER}");
    } 
}