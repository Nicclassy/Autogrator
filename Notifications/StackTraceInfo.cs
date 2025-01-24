using System.Diagnostics;

namespace Autogrator.Notifications;

public sealed record StackTraceInfo(string Method, string FileName, int LineNumber) {
    public static StackTraceInfo OfIndex(int index) {
        StackTrace stackTrace = new(fNeedFileInfo: true);
        StackFrame frame = stackTrace.GetFrame(index)!;

        string method = frame.GetMethod()!.Name;
        string filename = Path.GetFileName(frame.GetFileName())!;
        int lineNumber = frame.GetFileLineNumber();
        return new(method, filename, lineNumber);
    }
}