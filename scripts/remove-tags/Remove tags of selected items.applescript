# Summary: remove all tags from the selected items.
# "Items" in this context means projects, actions, or action groups.
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

# Display a dialog telling user of an error during script execution.
on display_error(msg)
	display dialog msg buttons {"OK"} with title my get_script_name() ¬
		with icon 0 giving up after 60
end display_error

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


# ~~~~ Main body ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

on run
	# We could launch OmniFocus if it's not running, but it doesn't make sense
	# for the user to invoke this script if OmniFocus is not running, so this
	# situation likely means something is wrong (e.g. running it accidentally).
	if not my is_app_running("OmniFocus") then
		my display_error("OmniFocus is not running.")
		return
	end if

	# If there's a usable selection, ask the user for confirmation.
	set selected_items to my get_selected_item_ids()
	set num_items to count of selected_items
	if num_items > 0 then
		set things to "selected item"
		if num_items > 1 then
			set things to num_items & " selected items"
		end if
		set msg to "Every tag on the " & things & " is about to be " ¬
			& "removed. This cannot be undone. Proceed?"
		display dialog msg buttons {"OK", "Cancel"} ¬
			with title my get_script_name() with icon 1 ¬
			default button 2 giving up after 60
		if button returned of result = "Cancel" or gave up of result then
			return
		end if
	else
		# The selection is empty or doesn't contain things that can have tags.
		return
	end if

	# Finally, let's try to do what we came here for.
	tell application "OmniFocus"
		try
			repeat with item_id in selected_items
				set this_item to my get_item(item_id)
				set current_tags to tags of this_item
				if count of current_tags > 0 then
					remove current_tags from tags of this_item
				end if
			end repeat
		on error err_msg number err_code
			set msg to err_msg & " (error code " & err_code & ")"
			my display_error(msg)
		end try
	end tell
end run
