# To-do

The following items have not yet been implemented for Autogrator and are potential things 
that could be implemented in the future:

* More control/configuration over downloading the senders list-also make the implementation 
specify whether or not the file is to be downloaded or local. Make the constructor load the list
* Parallelise some asynchronous methods by using `Task.WhenAll` in
async function calls that don't depend on each other
* More customisability over the notification emails file
* Implement a method `EnvVariableOrDefault` for greater flexibility in environment variables
* Make `SharePointClient`'s static factory method more accepting and using the booleans
from `AutogratorOptions`
* Handle the case for which the notification email environment variables are empty. If they
are empty but the `SendExceptionNotificationEmails` option is set to true, exit without an exception
* `IAllowedSenders` constructor must load the values from the file
* Improve response logging. The responses should be logged by the function
which obtains them, not by another function higher up in the call stack
* Configuration of the subject of automated emails

Some of these may never be implemeted, however the intention is to complete the important
items on the list.