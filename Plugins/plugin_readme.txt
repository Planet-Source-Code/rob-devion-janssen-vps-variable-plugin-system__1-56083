Plugin required functions...

DLL Name = VPS
Class name = Filename of DLL

fstrGetName (no args) return result should be a string giving the friendly name of the plugin
fstrGetFunction (no args) return result should be a string giving the function that initializes the plugin
fClose (no args) allows plugin to close itself before VPS terminates
registerparent (refObject) allows the plugin to be able to return calls.

Parent functions

Log(String) - logs to the log file and screen.
Logfile(String) - logs only to file
Logscreen(String) - logs only to screen
