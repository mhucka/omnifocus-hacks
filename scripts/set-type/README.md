# Set selected project or group type

As late as version 4.0, OmniFocus does not have a keyboard shortcut for changing the type of a project between sequential, parallel, and single-action. This curious omission means that you need to make your own. The script is my solution.

The script takes one argument, a project type, in the form of a string whose value should be one of `"parallel"`, `"sequential"`, or `"single action"` (the latter can be `"single"` for short). The script sets the types of all projects, groups, and/or tasks to the desired type, if possible. For items that cannot be of the given type (e.g., non-project items can't be set to single-action), it does nothing. For folders and tags, it also does nothing.

The type name can be passed to the script in one of two ways:

1. By invoking the script using the idiom

    ```AppleScript
    run script file SCRIPT_FILE with parameters {"TYPE"}
    ```

    where _SCRIPT_FILE_ is the full path to the compiled script and _TYPE_ is the name of the desired type; the latter should be one of the strings `"parallel"`, `"sequential"`, or `"single action"`.

2. By setting an environment variable named either `TypeName` or `KMVAR_TypeName` prior to running the script.

    (The purpose of checking `KMVAR_TypeName` is to support running this script from Keyboard Maestro using its "Execute AppleScript" action. You can use Keyboard Maestro's "Set variable to text" action before the "Execute AppleScript" action to set a variable named `TypeName` to the desired type, and then the value will be available to the script. Keyboard Maestro automatically defines environment variables for every variable you set, but the environment variable names it defines are always prefixed  with `KMVAR_`. This script checks both names, because outside of Keyboard Maestro, it's clearer to use a variable named `TypeName`.)

If the script can't find a type name using one of the methods above, it will display a dialog to ask for the type.
