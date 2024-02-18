# Summary: sort the tags of the currently-selected item(s).
# "Items" in this context means actions or action groups.
#
# Copyright 2024 Michael Hucka.
# License: MIT license – see file "LICENSE" in the project website.
# Website: https://github.com/mhucka/omnifocus-hacks

use AppleScript version "2.5"
use scripting additions


# ~~~~ Helping hands ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

# Return a nested list of tags in the same order as in the tags perspective.
# Warning: this is only designed to handle a 2-level tag hierarchy.
on get_tag_hierarchy()
	tell application "OmniFocus"
		tell front document
			set tag_hierarchy to {}
			repeat with _tag in every tag
				set subtags to every tag of _tag
				set end of tag_hierarchy to {_tag, subtags}
			end repeat
			return tag_hierarchy
		end tell
	end tell
end get_tag_hierarchy

# Return a list of all the names of the items in the given list.
on get_names(item_list)
	tell application "OmniFocus"
		set name_list to {}
		# Explicit loop because I couldn't get "name of every ..." to work.
		repeat with this_item in every item of item_list
			set end of name_list to name of this_item
		end repeat
		return name_list
	end tell
end get_names

# Take a list of tags & returns a list sorted according to the tag hierarchy.
on sort_tags(tag_list, tag_hierarchy)
	tell application "OmniFocus"
		set sorted_tags to {}
		repeat with subtree in tag_hierarchy
			set head_tag to item 1 of subtree
			repeat with given_tag in tag_list
				if name of given_tag = name of head_tag then
					set end of sorted_tags to head_tag
				end if
			end repeat
			repeat with subtag in (item 2 of subtree)
				repeat with given_tag in tag_list
					if name of given_tag = name of subtag then
						set end of sorted_tags to subtag
					end if
				end repeat
			end repeat
		end repeat
		return sorted_tags
	end tell
end sort_tags


# ~~~~ Main body ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

tell application "OmniFocus"
	try
		set tag_order to my get_tag_hierarchy()
		repeat with this_item in my get_selected_items()
			set current_tags to tags of this_item
			if count of current_tags > 1 then
				set sorted_tags to my sort_tags(current_tags, tag_order)
				if my get_names(sorted_tags) ≠ my get_names(current_tags) then
					remove current_tags from tags of this_item
					add sorted_tags to tags of this_item
				end if
			end if
 		end repeat
	on error err_msg number err_code
		set msg to err_msg & " (error code " & err_code & ")"
		display dialog msg buttons {"OK"} ¬
			with title "Script '" & my get_script_filename() & "'" ¬
			with icon 0 default button 1 giving up after 60
	end try
end tell
