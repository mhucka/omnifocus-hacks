# Summary: Given a tag name, toggle that tag in all selected items.
# "Items" in this context means actions or action groups.
#
# This script must be invoked with a single parameter, the name (or part of
# the name) of the tag to toggle on the selected items. This allows it to be
# invoked from the command line or other scripts using the idiom
#
#    run script file scripts_file with parameters {"tagname"}
#
# Copyright 2024 Michael Hucka.
# License: MIT license – see file "LICENSE" in the project website.
# Website: https://github.com/mhucka/keyboard-maestro-hacks

use AppleScript version "2.5" -- Yosemite (10.10) or later
use scripting additions


# Helper functions ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Return the file name of *this* script as a string.
on get_script_filename()
    tell application "System Events"
        set path_alias to path to me
		return name of path_alias
    end tell
end get_script_filename

# Return the currently selected tree(s) in OmniFocus.
on get_selection()
	tell application "OmniFocus"
		tell front window
			set tree_list to selected trees of content
			return every item of tree_list
		end tell
	end tell
end get_selection

# Return the currently selected actions, as a list.
on get_selected_items()
	tell application "OmniFocus"
		set item_list to {}
		repeat with selection in my get_selection()
			if class of selection is not in {tag, perspective, folder} then
				set end of item_list to value of selection
			end if
		end repeat
		return item_list
	end tell
end get_selected_items

# Remove the tag if it's found on the tag; otherwise, add the tag.
on toggle_tag(_task, _tag)
	tell application "OmniFocus"
		set tag_id to id of _tag
		repeat with existing_tag in tags of _task
			if (id of existing_tag) = tag_id then
				remove _tag from tags of _task
				return
			end if
		end repeat
		# If we didn't find the tag, add it.
		add _tag to tags of _task
	end tell
end toggle_tag


# Main body ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

on run {tag_name}
	if tag_name = ""
		tell application "OmniFocus"
			set answer to display dialog "Tag to toggle:" default answer "" ¬
				with title my get_script_filename() with icon 1 ¬
				buttons {"OK", "Cancel"} ¬
				default button "OK" cancel button "Cancel" ¬
				giving up after 30
			if button returned of answer = "Cancel" or tag_name = "" then
				return
			else
				set tag_name to text returned of result
			end if
		end tell
	end if

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

	# Now do what we came here to do.
	tell application "OmniFocus"
		# Get the actual tag object for the named tag.
		tell front document
			set found_tag to (complete tag_name as tag)
			if found_tag ≠ {} then
				set tag_id to id of item 1 of found_tag
				set the_tag to first flattened tag whose id is tag_id
			else
				set msg to "Could not find a tag with the text '" & ¬
					tag_name & "' in its name in OmniFocus."
				display dialog msg buttons {"OK"} ¬
					with title "Script '" & my get_script_filename() & "'" ¬
					with icon 0 default button 1 giving up after 60
				return
			end if
		end tell
		
		# Iterate over the selected items and toggle the tag.
		repeat with this_item in my get_selected_items()
			my toggle_tag(this_item, the_tag)
		end repeat
	end tell
end run
