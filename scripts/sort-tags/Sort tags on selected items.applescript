# Summary: sort the tags of the currently-selected item(s).
#
# Copyright 2024 Michael Hucka.
# License: MIT license – see file "LICENSE" in the project website.
# Website: https://github.com/mhucka/omnifocus-hacks

use AppleScript version "2.5"
use scripting additions


# ~~~~ Helping hands ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

on get_selection()
	tell application "OmniFocus"
		tell front window
			set tree_list to selected trees of content
			return every item of tree_list
		end tell
	end tell
end get_selection

on get_selected_actions()
	tell application "OmniFocus"
		set actions to {}
		repeat with selection in my get_selection()
			if class of selection is not in {tag, perspective, folder} then
				set end of actions to value of selection
			end if
		end repeat
		return actions
	end tell
end get_selected_actions

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

# Takes a list of tags & returns a list sorted according to the tag hierarchy.
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
	set tag_order to my get_tag_hierarchy()
	tell front document
		repeat with action in my get_selected_actions()
			set current_tags to tags of action
			set sorted_tags to my sort_tags(current_tags, tag_order)
			if sorted_tags ≠ current_tags then
				remove current_tags from tags of action
				add sorted_tags to tags of action
			end if
 		end repeat
	end tell
end tell
