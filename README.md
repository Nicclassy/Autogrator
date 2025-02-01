# Autogrator
Autogrator is a .NET application that automatically uploads emails 
from Microsoft Outlook to Microsoft SharePoint.

## About
The process which the application undertakes to achieve its objective 
can be divided broadly into the following steps:
login to (classic) Outlook; process emails when they are received and optionally, restrict certain senders;
export emails to a desired format (by default, PDFs); 
create—and check for the existence of—folders for the emails to be stored in 
and finally; upload the emails to SharePoint. 
Autogrator provides customisability at every stage of this process and uses Serilog to log the actions
it performs.

The name 'Autogrator' is a truncated, denominalised portmanteau of 'automatic' and 'integration'.
Autogrator is currently being hosted on Microsoft Azure VM(s) as you are reading this.

# Setup
The setup process is rather lengthy due to the tighly-coupled nature of Microsoft's applications, services and Windows itself. 
Containerisation of this application via Docker or other software is obviously infeasible due to the basic requirement
of having Microsoft Outlook installed and authenticated. \
Therefore, the first pre-requisite of running Autogrator is that you have
a Windows VM to use.

## Prerequisites
* .NET 9.0
* A Windows VM if you plan to use Autogrator automatically (you do not need one for its setup, however)
* A user with sufficient administrative priviledges in Microsoft Entra admin center
to grant appropriate application permissions for your registered application
* (Optional) `msbuild` if you want to build Autogrator manually

## Online Setup

1. [Create a new app registration](https://learn.microsoft.com/en-us/graph/toolkit/get-started/add-aad-app-registration)
at https://entra.microsoft.com. You do not need to include a Redirect URI

2. Once you have created your application, note its Client ID (which is also known as its Application ID) and
its Tenant ID (also known as its Directory ID). You will need these values later to store in the `.env` file

3. Create a new 'Client Secret' for your app registration by navigating to the 'Certificates & secrets' menu, and note the value
as you will need it later

4. Navigate to your application's 'API Permissions' and add the application permission `Sites.ReadWrite.All`. 
Be wary that you will need sufficient administrative priviledges to grant this permission.

## Local Setup

1. Clone this repository using the command
```powershell
git clone "https://github.com/Nicclassy/Autogrator"
```

2. Install [Microsoft Office 365](https://www.microsoft.com/en-us/microsoft-365/download-office) 
and login to Outlook

3. Copy `.env.example`, rename it to `.env` and fill out the missing values.
The next two steps provides context with regard to 
`AG_EMAIL_CONTENT_PATH` and `AG_ALLOWED_SENDERS_*` but are optional.
For the values pertaining to SharePoint (e.g. site paths/drive names),
see [Obtaining SharePoint Infromation](#obtaining-sharepoint-information).
You can perform this step by using the following commands in PowerShell:
```powershell
Copy-Item .env.example .env
# Alternatively, you can do: copy .env.example .env
# if you are using Command Prompt
your_favourite_editor .env
```

4. (Optional) For the variable `AG_EMAIL_CONTENT_PATH`, provide a file path—whose contents will be read by `File.ReadAllText`. 
If you would like to use the automated notification emails feature for more detailed diagnostics,
you can use the following holes in the text of the file (as the file contents are just a 
[message template](https://messagetemplates.org)):
    - `{LineNumber}` - The line number on which the exception is thrown
    - `{FileName}` - The name of the file from which the exception was thrown
    - `{Method}` - The name of the method from which the exception was thrown
    - `{ExceptionType}` - The type of exception which was thrown
    - `{TimeStamp}` - The timestamp of the thrown exception, which is determined immediately

    An example file can be shown below:
    ```
    This is an automatically generated email sent from Autogrator.

    Autogrator crashed at {TimeStamp} on line {LineNumber} in file {FileName} in method {Method} with an exception of type {ExceptionType}.
    Please see the attached log file for more details.
    ```

5. (Optional) For restricting/allowing email senders based on predefined rules, you can provide your own 
implementation of `IAllowedSenders` for use with `Autogrator.Builder`. The corresponding environment
variables are prefixed with `AG_ALLOWED_SENDERS`. The file is downloaded from the SharePoint you specify.
If you do not provide an implementation of `IAllowedSenders` then `AllEmailSendersAllowed` will be used 
(which allows all email addresses that are valid under 
[RFC 5322](https://datatracker.ietf.org/doc/html/rfc5322#section-3.4)).
You can use your on implementation of `IAllowedSenders` with `Autogrator.Builder` like this:
```csharp
using Autogrator;

Autogrator autogrator = new Autogrator.Builder()
    .WithAllowedSenders(new YourAllowedSenders())
    .Build();
```

## Running Autogrator
There are three ways to run Autogrator:

a. Use Autogrator in your code. Configure the instance of `Autogrator` using `Autogrator.Builder`
and `AutogratorOptions`. For example:
```csharp
using Autogrator;

Autogrator autogrator = new Autogrator.Builder()
    .WithAllowedSenders(new ExcelAllowedSenders())
    .WithEmailFileNameFormatter(mailItem => $"{mailItem.SenderName} {mailItem.CreationTime}")
    .WithOptions(new AutogratorOptions {
        OverwriteDownloads = true,
        SendExceptionNotificationEmails = true
    })
    .Build();
autogrator.Run();
```

b. Use Visual Studio Code to run Autogrator. This merely requires you to click the 'Start Without Debugging' button.

c. Build and run Autogrator manually using the following commands (this requires `msbuild` to be installed):
```powershell
dotnet restore
msbuild /t:Build /p:Configuration=Release
start ".\Autogrator\bin\Release\net9.0\Autogrator.exe"
```

## Obtaining SharePoint Information
In the fictional SharePoint URL
```html
https://autogrator.sharepoint.com/sites/Autogrator/Documentation
```
the hostname is `autogrator.sharepoint.com`, 
the site path is `/sites/Autogrator` and the drive name is `Documentation`. Experiment with
[Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) 
if you are unfamiliar with how the Microsoft Graph API works and would like to better understand how to use it.
