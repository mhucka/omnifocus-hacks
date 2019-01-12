-- ============================================================================
-- @file    of-toggle-tag
-- @brief   Add or remove one or more tags from tasks
-- @author  Michael Hucka <mhucka@caltech.edu>
-- @license Please see the file LICENSE in the parent directory
-- @repo    https://github.com/mhucka/omnifocus-hacks
-- ============================================================================

-- Global variables and constants.
-- ............................................................................

global tagsList
set tagsList to {"today", "soon"}

-- Utility functions.
-- ............................................................................

on askUserForTags()
	tell application "System Events"
		activate
		set selectedTags to {choose from list tagsList with title "Toggle Tags 🏷" with prompt "Tags to toggle:" with multiple selections allowed}
	end tell
	-- We get back a list of lists. I don't know why. Return the inner one.
	return item 1 of selectedTags
end askUserForTags

-- The following is based on code posted by user "RosemaryOrchard" here:
-- https://talk.automators.fm/t/running-into-problems-adding-a-tag-in-omnifocus-3-with-an-applescript/2227/2

on tagObjectFromOF(theTagName)
	tell application "OmniFocus"
		tell default document
			return the first flattened tag where its name = theTagName
		end tell
	end tell
end tagObjectFromOF


-- The following is based in part on code posted by user "hammer" in Oct. 2018
-- https://discourse.omnigroup.com/t/omnifocus-3-script-to-remove-today-tag-and-mark-complete/42541/2

on toggleTag(theTask, theTagName)
	tell application "OmniFocus"
		-- First look for the tag in the task's tags & remove it if it's there.
		set currTags to name of tags of theTask
		repeat with x from 1 to count of currTags
			set currTagName to item x of currTags
			if currTagName = theTagName then
				set currTag to (first tag of theTask whose name is item x of currTags)
				remove currTag from tags of theTask
				return
			end if
		end repeat

		-- If we get here, we didn't find the tag. Add it.
		set theTag to my tagObjectFromOF(theTagName)
		add theTag to tags of theTask
	end tell
end toggleTag


-- Main body.
-- ............................................................................

set listOfTagsToToggle to my askUserForTags()

tell application "OmniFocus"
	activate
	delay 0.2

	tell front window
		set selectedTrees to selected trees of content
		set selectedTasks to every item of selectedTrees
	end tell
	tell front document
		repeat with i from 1 to count of selectedTasks
			set theTask to the value of item i of selectedTasks
			repeat with i from 1 to count of listOfTagsToToggle
				set theTagName to item i of listOfTagsToToggle
				my toggleTag(theTask, theTagName)
			end repeat
		end repeat
	end tell
end tell
