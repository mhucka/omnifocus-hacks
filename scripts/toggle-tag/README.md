# Toggle a given tag on selected items

This script tags the name of a tag, and toggles the tag on the selected items in OmniFocus. If the tag is present on some items, it's removed from those items; if it's not present, it's added to items.

The script is designed to be usable standalone as well as part of Keyboard Maestro actions (which is how I use it).  To invoke it in Keyboard Maestro, first copy the compiled script to `~/Library/Scripts/Applications/OmniFocus"`, then use the [Execute AppleScript](https://wiki.keyboardmaestro.com/action/Execute_an_AppleScript) action.

```applescript
# Set this to the tag name ↓
set script_parameters to {"eventually"}

# Set this to the name of this script (without path) ↓
set script_filename to "Toggle given tag on selected items.scpt"

# The remaining lines are generic invocation code should be left as-is.
set lib to (path to library folder from user domain as text)
set script_file to lib & "Scripts:Applications:OmniFocus:" & script_filename
run script file script_file with parameters script_parameters
```

Keyboard Maestro has a mechanism for passing arguments to scripts, but it requires that the script itself has knowledge about how to access the parameter in KM, and it also means that the KM macro and the script become linked which requires more maintenance. (E.g., if you change the name of the variable, you have to change it in both places. Plus, users of this script have to be informed about how to do this, and so on.)

The advantage of doing this is that the script does not rely on any knowledge about KM. The information it needs is passed as a parameter.
