(*
AppleScript for InDesign to migrate GREP styles from a user-chosen paragraph style in a source doc to a user-chosen paragraph style in a destination document. If 2 or more documents are open, user chooses source and destination from list. Otherwise, user chooses from a file dialog.

Copyright ©2013 by Rick Gordon, with permission for public use. Please credit me if you mention this script. Thank you.
*)

tell application "Adobe InDesign CS5.5"
  set vApp to it
	set vDocCount to count documents
	if vDocCount ≥ 2 then
		set vDocNameList to name of documents
		tell me to set vSourceDocName to item 1 of (choose from list vDocNameList with prompt "Choose Source Document." default items {item 1 of vDocNameList})
		set DestDocNameList to {}
		repeat with vItem in vDocNameList
			if vItem as string is not equal to vSourceDocName then
				set end of DestDocNameList to vItem
			end if
		end repeat
		tell me to set vDestDocName to item 1 of (choose from list DestDocNameList with prompt "Choose Destination Document." default items {item 1 of DestDocNameList})
		set vSourceDoc to document vSourceDocName
		set vDestDoc to document vDestDocName
	else
		tell me
			set vSourceDocAlias to choose file with prompt "Choose Source File."
			repeat
				set vDestDocAlias to choose file with prompt "Choose Destination File."
				if vDestDocAlias is equal to vSourceDocAlias then
					display alert "Source and destination must be different files."
				else
					exit repeat
				end if
			end repeat
		end tell
		open vDestDocAlias
		set vDestDoc to document 1
		open vSourceDocAlias
		set vSourceDoc to document 1
	end if
	
	tell vSourceDoc
		set vsourceParaStyleList to object reference of items of all paragraph styles
		set vSourceNameList to {}
		repeat with vItem in vsourceParaStyleList
			tell vItem
				if properties of nested grep styles is not {} then
					if class of parent is paragraph style group then
						set end of vSourceNameList to name of parent & ": " & name
					else
						set end of vSourceNameList to name
					end if
				end if
			end tell
		end repeat
		tell me to set vChosenItem to item 1 of (choose from list vSourceNameList with prompt "Choose Source Paragraph Style.")
		set my text item delimiters to {": "}
		set vChosenName to last item of text items of vChosenItem
		set my text item delimiters to {""}
		set vCount to count items in vsourceParaStyleList
		repeat with i from 1 to vCount
			if vChosenName is equal to ((name of item i in vsourceParaStyleList) as string) then
				set vsourceParaStyleRef to object reference of item i in vsourceParaStyleList
			end if
		end repeat
		set vGrepExpression to grep expression of nested grep styles of vsourceParaStyleRef
		tell me to set vGrepList to (choose from list vGrepExpression with prompt ¬
			"Choose one or more from the  source GREP Style." default items vGrepExpression with multiple selections allowed)
		set vSourceGrepRefList to {}
		set vCharStyleRefList to {}
		tell vsourceParaStyleRef
			repeat with i from 1 to count items in vGrepList
				set vGrepStyleRef to (object reference of nested grep styles where grep expression is equal to item i of vGrepList)
				set end of vSourceGrepRefList to vGrepStyleRef
				set end of vCharStyleRefList to object reference of applied character style of vGrepStyleRef
			end repeat
		end tell
	end tell
	
	tell vDestDoc
		set vDestParaStyleList to object reference of items 2 thru -1 of all paragraph styles
		set vDestNameList to {}
		repeat with vItem in vDestParaStyleList
			tell vItem
				if class of parent is paragraph style group then
					set end of vDestNameList to name of parent & ": " & name
				else
					set end of vDestNameList to name
				end if
			end tell
		end repeat
		tell me to set vChosenItem to item 1 of (choose from list vDestNameList with prompt "Choose destination paragraph style to add GREP styles into.")
		
		tell paragraph style vChosenItem
			repeat with i from 1 to (count items in vGrepList)
				set vCurrentSourceGrepRef to item i of vSourceGrepRefList
				set vCharStyleName to name of applied character style of vCurrentSourceGrepRef
				set vDoesNotAlreadyExist to ((nested grep styles where grep expression is equal to grep expression of vCurrentSourceGrepRef) is equal to {}) is true
				if vDoesNotAlreadyExist then
					set vDestAppliedCharStyle to (items of all character styles of vDestDoc where name is equal to vCharStyleName)
					if length of vDestAppliedCharStyle > 0 then
						set vDestAppliedCharStyle to item 1 of vDestAppliedCharStyle
					else
						tell vSourceDoc
							set vSourceTempFrame to make new text frame with properties {fill color:"Black", fill tint:20, geometric bounds:{0, "-6p", "4p", "-2p"}, nonprinting:true, label:"TEMP"}
							set properties of insertion point 1 of parent story of vSourceTempFrame to {contents:"X", applied paragraph style:paragraph style "[No Paragraph Style]", applied character style:applied character style of vCurrentSourceGrepRef}
							set vSourceWindow to layout window 1
							bring to front vSourceWindow
							select vSourceTempFrame
							tell vApp to copy
						end tell
						tell vDestDoc
							set vDestWindow to layout window 1
							bring to front vDestWindow
							tell vApp to paste
							set vDestTempFrame to (some text frame where label is "TEMP")
						end tell
						set vDestAppliedCharStyle to applied character style of character 1 of parent story of vDestTempFrame
						tell vSourceDoc to delete vSourceTempFrame
						tell vDestDoc to delete vDestTempFrame
						--duplicate vCurrentSourceGrepRef to end of vDestDoc --NOPE
						--tell vDestDoc to set vDestAppliedCharStyle to (make new character style at end with properties (properties of applied character style of vCurrentSourceGrepRef)) --NOPE
						--set vDestAppliedCharStyle to applied character style of vCurrentSourceGrepRef --NEED TO GET CHAR STYLE INTO DOC FIRST
					end if
					--return items of vDestCharStyleNames
					make new nested grep style at end with properties {applied character style:vDestAppliedCharStyle, grep expression:grep expression of vCurrentSourceGrepRef}
				end if
			end repeat
		end tell
	end tell
end tell
