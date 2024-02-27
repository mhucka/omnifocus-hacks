# Summary: set the project type of the selected projects.
#
# Copyright 2024 Michael Hucka.
# License: MIT license – see file "LICENSE" in the project website.
# Website: https://github.com/mhucka/omnifocus-hacks

use AppleScript version "2.5"
use scripting additions

# ~~~~ Helping hands ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Return the file name of *this* script as a string.
on get_script_filename()
	local path_alias
    tell application "System Events"
        set path_alias to path to me
		return name of path_alias
    end tell
end get_script_filename

# Return a list of item id's for the currently-selected items in OmniFocus.
on get_selected_item_ids()
	local item_links, item_ids, orig_delims

	# There's a bug in the AppleScript interface of OmniFocus that has existed
	# since at least 2008. The bug is this. If you select a task or group in
	# the content window, then select the parent project in the sidebar, there
	# is nothing in the available AppleScript objects that can be used to
	# determine whether the actual highlighted selection is (a) the parent
	# project in the sidebar or (b) the items in the content window. The
	# property values of elements in "selected trees of" for both "content"
	# and "sidebar" are absolutely the same in both cases. You can use the
	# OmniFocus menu items "Go to outline" and "Go to "sidebar", go back and
	# forth, and the objects returned by OmniFocus's AppleScript interface
	# remain same each time. Thus, you can't tell what the user has truly
	# selected. It's an exasperating problem because, obviously, OmniFocus
	# itself knows what the user has selected. The following hack is a
	# workaround. It gets the selected items using OmniFocus's menu item "Copy
	# as Link", which doesn't suffer from the same problem. I don't like this,
	# and I spent hours looking for a better way. This is the best I found.

	set item_links to {}
	tell application "System Events"
		set my_app to bundle identifier of first process whose frontmost is true
		tell application process "OmniFocus"
			# The menu item click won't take effect unless you set frontmost.
			set frontmost to true
			tell menu bar item "Edit" of menu bar 1
				click menu item "Copy as Link" of menu 1
				delay 0.1
			end tell
			set item_links to paragraphs of (the clipboard as text)
		end tell
		activate application my_app
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

# Given an id or an item link, return the object (task, group, or project).
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
on set_item_type(item_obj, desired_type)
	tell application "OmniFocus"
		set is_sequential to (sequential of item_obj)
		# Projects can be single-action, but tasks & groups cannot. It would
		# be annoying to show an error dialog if the user tried to set a
		# task or group to single-action, so we just quietly ignore it.
		try
			set is_single to (singleton action holder of item_obj)
		end try

		if desired_type = "single" then
			if is_sequential then
				set sequential of item_obj to false
			end if
			if not is_single then
				try
					set item_obj's singleton action holder to true
				end try
			end if
		else if desired_type = "sequential" then
			if is_single then
				try
					set singleton action holder of item_obj to false
				end try
			end if
			if not is_sequential then
				set sequential of item_obj to true
			end if
		else if desired_type = "parallel" then
			# Parallel is the default if it's not sequential or single-action.
			if is_single then
				try
					set singleton action holder of item_obj to false
				end try
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
	tell application "System Events"
		if count of (every process whose name is "OmniFocus") = 0 then
			display dialog "OmniFocus is not running." buttons {"OK"} ¬
				with title "Script '" & my get_script_filename() & "'" ¬
				with icon 0 default button 1 giving up after 60
			return
		end if
	end tell

	# If the tag name was not passed as an argument, check the environment.
	if (type_name as string) = "" then
		set type_name to system attribute "KMVAR_TagName"
		if type_name = "" then
			set type_name to system attribute "TagName"
		end if
	end if

	# If we still don't have a tag name, ask the user for one.
	if type_name = "" then
		tell application "OmniFocus"
			set answer to display dialog "Desired type:" default answer "" ¬
				with title my get_script_filename() with icon 1 ¬
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
			with title "Script '" & my get_script_filename() & "'" ¬
			with icon 0 default button 1 giving up after 60
		return
	end if

	# Normalize the type name.
	if type_name is in {"single", "single action", "single-action"} then
		set type_name to "single"
	end if

	# Finally, let's try to do what we came here for.
	repeat with item_id in my get_selected_item_ids()
		my set_item_type(my get_item(item_id), type_name)
	end repeat
end run
