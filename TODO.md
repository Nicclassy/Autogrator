# To-do

The following items have not yet been implemented for Autogrator and is planned to be
implemented in the future:

* Scripts `run.cmd` and `run.ps1` for quickly running the program. Also `run.sh` if supported
* More control/configuration over downloading the senders list-also make the implementation 
specify whether or not the file is to be downloaded or local. Make the constructor load the list
* More customisability over the notification emails file
* Implement a method `EnvVariableOrDefault` for greater flexibility in environment variables
* Make `SharePointClient`'s static factory method more accepting and using the booleans
from `AutogratorOptions`
* Handle the case for which the notification email environment variables are empty. If they
are empty but the `SendExceptionNotificationEmails` option is set to true, exit without an exception
* `GraphHttpClient.LogFailureAndThrow` should follow correct message templating conventions
* `EmailReceiver.LogRejectedSenders` should be configured in `AutogratorOptions`
* Configuration of the subject of automated emails
* Testing if Autogrator runs on MacOS

Some of these may never be implemeted, however the intention is to complete the important
items on the list.