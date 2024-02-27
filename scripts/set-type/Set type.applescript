# Summary: set the project type of the selected project, task, or group.
#
# This script takes one argument, a project type, in the form of a string
# whose value should be one of "parallel", "sequential", or "single action"
# (the latter can be "single" for short). The script sets the types of all
# projects, groups, and/or tasks to the desired type, if possible. For items
# that cannot be of the given type (e.g., non-project items can't be set to
# single-action), it does nothing. For folders and tags, it also does nothing.
#
# The type name can be passed to the script in one of two ways:
#
# 1) By invoking the script using the idiom
#        run script file SCRIPTFILE with parameters {"TYPE"}
#    where SCRIPTFILE is the full path to the compiled script and TYPE is
#    the name of type ("parallel", "sequential", or "single action").
#
# 2) By setting an environment variable named either "TypeName" or
#    "KMVAR_TypeName" prior to running the script. The "KMVAR_TypeName" form
#    is available to support running this script from Keyboard Maestro using
#    its "Execute AppleScript" action. Use K.M.'s "Set variable to text"
#    action before the "Execute AppleScript" action.
#
# If the script can't find a type name using one of the methods above, it
# will display a dialog to ask for the type.
#
# Copyright 2024 Michael Hucka.
# License: MIT license – see file "LICENSE" in the project website.
# Website: https://github.com/mhucka/omnifocus-hacks

use AppleScript version "2.5"
use scripting additions


# ~~~~ Helping hands ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Performance note: this script runs noticeably slower if "use framework
# Foundation" is put at the top level than if it's used inside the (one)
# handler it's needed below.

# Return the given file name without its file name extension, if any.
on remove_ext(file_name)
	script wrapperScript
		property ca: a reference to current application
		use framework "Foundation"
		on remove_ext(file_name)
			set u to ca's NSURL's fileURLWithPath:file_name
			return u's URLByDeletingPathExtension()'s lastPathComponent() as text
		end remove_ext
	end script
	return wrapperScript's remove_ext(file_name)
end remove_ext

# Return the file name of this script as a string, minus the extension.
on get_script_name()
    tell application "System Events"
        set path_alias to path to me
		set file_name to name of path_alias
		return my remove_ext(file_name)
    end tell
end get_script_name

# Return true if the named application is running.
on is_app_running(app_name)
	tell application "System Events"
		return (count of (every process whose name is "OmniFocus")) > 0
	end tell
end is_app_running

# Return a list of item id's for the currently-selected items in OmniFocus.
on get_selected_item_ids()
	local item_links, item_ids, orig_delims

	# A bug in the OmniFocus AppleScript interface that has existed since at
	# least 2008 is still present in version 4.0.5 today (in 2024). If you
	# select a task in the content window, nothing in the AppleScript objects
	# can be used to determine whether the actual highlighted selection is (a)
	# the task in the content window or (b) the parent project in the sidebar.
	# If you use "Go to sidebar" (from the OmniFocus user interface View
	# menu), get both the "selected of sidebar" and "selected of content"
	# objects via AppleScript, then use "Go to outline", get the AppleScript
	# objects, and compare them, you'll find that the property values of the
	# relevant objects are the same. The property that *should* distinguish
	# them (a Boolean property named "selected" -- a confusing name in this
	# context) always has the value false (as reported by user "davidamis" in
	# 2008-05-29 at http://forums.omnigroup.com/showthread.php?p=37446). Thus,
	# you can't tell what the user has truly selected. It's an exasperating
	# problem because, obviously, OmniFocus itself knows what the user has
	# selected. The following hack is a workaround. It gets the selected items
	# using OmniFocus's menu item "Copy as Link", because it returns the true
	# selections. I don't like this GUI scripting, but I spent many hours
	# looking for a better way and this is the best alternative I found.

	set item_links to {}
	tell application "System Events"
		tell application process "OmniFocus"
			# The menu item click won't take effect unless you set frontmost.
			set frontmost to true
			tell menu bar item "Edit" of menu bar 1
				click menu item "Copy as Link" of menu 1
				delay 0.1
			end tell
			set item_links to paragraphs of (the clipboard as text)
		end tell
	end tell

	# Extract the item id's from the item links.
	set item_ids to {}
	set orig_delims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ["/"]	
	repeat with this_link in item_links
		if contents of this_link is not "" then
			set end of item_ids to last text item of this_link
		end if
	end repeat
	set AppleScript's text item delimiters to orig_delims

	# Return what we got.
	return item_ids
end get_selected_item_ids

# Given an id or an item link for a task or project, return the object.
on get_item(item_id)
	tell application "OmniFocus"
		tell default document
			set obj to ""
			try
				set obj to first flattened project where its id is item_id
			end try
			if obj = "" then
				try
					set obj to first flattened task where its id is item_id
				end try
			end if
			return obj
		end tell
	end tell
end get_item

# Set the type (parallel/sequential/single-action) of the item.
on set_item_type(item_id, desired_type)
	set item_obj to my get_item(item_id)
	tell application "OmniFocus"
		set is_sequential to (sequential of item_obj)
		set is_single to false
		set can_be_single to false
		# It's harder to test what kind of object we have, than to do this.
		try
			set is_single to (singleton action holder of item_obj)
			set can_be_single to true
		end try

		if desired_type = "single" and can_be_single then
			if is_sequential then
				set sequential of item_obj to false
			end if
			if not is_single then
				set item_obj's singleton action holder to true
			end if
		else if desired_type = "sequential" then
			if can_be_single and is_single then
				set singleton action holder of item_obj to false
			end if
			if not is_sequential then
				set sequential of item_obj to true
			end if
		else if desired_type = "parallel" then
			# Parallel is the default if it's not sequential or single-action.
			if can_be_single and is_single then
				set singleton action holder of item_obj to false
			end if
			if is_sequential then
				set sequential of item_obj to false
			end if
		end if
	end tell
end set_item_type


# ~~~~ Main body ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

on run type_name
	# We could launch OmniFocus if it's not running, but it doesn't make sense
	# for the user to invoke this script if OmniFocus is not running, so this
	# situation likely means something is wrong (e.g. running it accidentally).
	if not my is_app_running("OmniFocus") then
		display dialog "OmniFocus is not running." buttons {"OK"} ¬
			with title "Script '" & my get_script_name() & "'" ¬
			with icon 0 default button 1 giving up after 60
		return
	end if

	# If the type name was not passed as an argument, check environment vars.
	if (type_name as string) = "" then
		set type_name to system attribute "KMVAR_TypeName"
		if type_name = "" then
			set type_name to system attribute "TypeName"
		end if
	end if

	# If we still don't have a type name, ask the user for one.
	if type_name = "" then
		tell application "OmniFocus"
			set answer to display dialog "Desired type:" default answer "" ¬
				with title my get_script_name() with icon 1 ¬
				buttons {"OK", "Cancel"} ¬
				default button "OK" cancel button "Cancel" ¬
				giving up after 30
			if button returned of answer = "Cancel" or type_name = "" then
				return
			else if gave up of result then
				return
			else
				set type_name to text returned of result
			end if
		end tell
	end if

	# Make sure we got a valid type name.
	if not type_name is in {"parallel", "sequential", "singleton", ¬
							"single", "single action", "single-action"} then
		set msg to "Unable to interpret \"" & type_name & "\" as a known " ¬
			& " type. Please use one of the names \"parallel\", " ¬
			& "\"sequential\", or \"single action\" (or just \"single\")."
		display dialog msg buttons {"OK"} ¬
			with title "Script '" & my get_script_name() & "'" ¬
			with icon 0 default button 1 giving up after 60
		return
	end if

	# Normalize the type name.
	if type_name is in {"single", "single action", "single-action"} then
		set type_name to "single"
	end if

	# Finally, let's try to do what we came here for.
	try
		repeat with this_id in my get_selected_item_ids()
			my set_item_type(this_id, type_name)
		end repeat
	on error err_msg number err_code
		set msg to err_msg & " (error code " & err_code & ")"
		display dialog msg buttons {"OK"} ¬
			with title "Script '" & my get_script_name() & "'" ¬
			with icon 0 default button 1 giving up after 60
	end try
end run
