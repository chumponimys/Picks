-- ExcelFuncts.applescript
-- Picks

-- Created by Josh Fletcher on 7/29/12.
-- Copyright 2012 Ari Cohen. All rights reserved.


script ExcelFuncts
	property parent : class "NSObject"
	property gameNumbers : {16, 16, 16, 15, 14, 14, 13, 14, 14, 14, 14, 16, 16, 16, 16, 16, 16}
	property alphab : {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	property critLines : {5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35}
	property ASTID : AppleScript's text item delimiters
	property stopper : "Go"
	
	
	on exportMe_(sender)
		delay 2
		log "TEST"
		delay 2
	end exportMe_
	
	on exportAllForms_(myInfo)
		try
			set PicksPath to ((item 1 of myInfo) as string) as alias
			set thePicksPath to POSIX path of (PicksPath)
			set ASTID to AppleScript's text item delimiters
			set TheWeek to (((item 2 of myInfo) as string) as number)
			if TheWeek < 10 then set TheWeek to "0" & TheWeek
			set TheWeek to TheWeek as string
			set progBar to item 3 of myInfo
			set theStat to item 6 of myInfo
			
			tell application "Mail" to set TheSubjects to (sender of every message in mailbox ("w" & TheWeek) of mailbox "FP14")
			set theCounter to 0
			
			progBar's setIndeterminate_(false)
			progBar's setMaxValue_(count TheSubjects)
			progBar's setDoubleValue_(0)
			theStat's setStringValue_("Exporting Forms From Mail...")
			doEventFetch()
			
			repeat with s in TheSubjects
				set theCounter to theCounter + 1
				tell application "Mail"
					set theContent to ((content of message theCounter in mailbox ("w" & TheWeek) of mailbox "FP14") as string)
					set theDate to ((date received of message theCounter in mailbox ("w" & TheWeek) of mailbox "FP14") as string)
				end tell
				set theTime to (do shell script "echo " & quoted form of theDate & " | cut -f3 -d\",\" | cut -c7-")
				set thePlaceHolder to (do shell script "echo " & quoted form of theTime & " | cut -f1 -d\" \"")
				set theTime to (thePlaceHolder & "_" & (do shell script "echo " & quoted form of theTime & " | cut -f2 -d\" \""))
				set AppleScript's text item delimiters to ":"
				set theTime to text items of theTime
				set AppleScript's text item delimiters to "."
				set theTime to theTime as string
				set AppleScript's text item delimiters to ASTID
				set theDate to (do shell script "echo " & quoted form of theDate & " | cut -f2 -d\",\" | cut -c2-")
				set thePlaceHolder to theDate
				set theDate to (do shell script "echo " & quoted form of theDate & " | cut -f1 -d\" \"")
				set theDate to (theDate & "_" & (do shell script "echo " & quoted form of thePlaceHolder & " | cut -f2 -d\" \""))
				set theDate to theDate & "_" & theTime
				set NumberName to ((do shell script "echo \"" & theContent & "\" | sed -n '2p' | cut -c1-4"))
				do shell script "echo \"" & theContent & "\" > " & thePicksPath & NumberName & "_" & theDate & "_" & theCounter & ".txt"
				progBar's incrementBy_(1)
			end repeat
		on error theError
			display dialog "MAIL ERROR: " & theError
			error theError
		end try
	end exportAllForms_
	
	on exportSelectForms_(sender)
	end exportSelectForms_
	
	on progWinDidEnd_returnCode_contextInfo_(theSheet, returnCode, unUsed)
		tell theSheet to orderOut_(me) -- now close the sheet
	end progWinDidEnd_returnCode_contextInfo_
	
	on EnterData_(myInfo)
		try
			set APath to (((item 1 of myInfo) as string) as alias)
			set PPath to POSIX path of (APath)
			set TheWeek to ((item 2 of myInfo) as string) as number
			set NumGames to item TheWeek of gameNumbers
			set progBar to item 3 of myInfo
			set theStat to item 6 of myInfo
			set theIDs to {}
			set IDNums to {}
			set dupArray to {}
			
			set theParagraphs to (paragraphs of (do shell script "ls " & quoted form of PPath))
			
			progBar's setIndeterminate_(false)
			progBar's setMaxValue_(((count theParagraphs) * 5) + 50)
			progBar's setDoubleValue_(0)
			theStat's setStringValue_("Reading Files...")
			doEventFetch()
			set AppleScript's text item delimiters to "_"
			repeat with z in theParagraphs
				--doEventFetch()
				--if (my stopper = "stop") then error "STOP"
				set thenum to (text item 1 of z)
				set cont to true
				try
					set thenum to ((thenum as number) as string) --Check if correct file name type
				on error
					set cont to false
					tell application "Finder" to set label index of (((APath as string) & z) as alias) to 6
					progBar's incrementBy_(5)
				end try
				if cont then
					if IDNums contains thenum then
						tell application "Finder" to set label index of (((APath as string) & z) as alias) to 4
						set end of dupArray to thenum
						progBar's incrementBy_(5)
					else
						set end of IDNums to thenum
						set end of theIDs to {z as string, thenum}
						progBar's incrementBy_(1)
					end if
				end if
			end repeat
			set AppleScript's text item delimiters to ASTID
			set theValues to {}
			set thenum to 2
			theStat's setStringValue_("Reading All ID's From Excel Sheet...")
			doEventFetch()
			repeat
				doEventFetch()
				--if (my stopper = "stop") then error "STOP"
				set thenum to 1 + thenum
				set c to "A" & thenum
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set val to value of cell c --Generate array theValues of every Number ID in Excel sheet for cross reference
				set end of theValues to val
				if val = "stop" then exit repeat
			end repeat
			progBar's incrementBy_(50)
			
			set counter to 0
			theStat's setStringValue_("Checking Files For Duplicates and Invalid ID's...")
			doEventFetch()
			repeat with x in theIDs
				--doEventFetch()
				--if (my stopper = "stop") then error "STOP"
				if (dupArray contains (item 2 of x)) or (theValues does not contain (item 2 of x as number)) then --If file is duplicate or conatins invalid Number ID
					if (theValues does not contain (item 2 of x as number)) then set theMod to 2
					if (dupArray contains (item 2 of x)) then set theMod to 4
					log theMod
					set myPath to (((APath as string) & (item 1 of x)) as string) as alias
					--log myPath as string
					--tell application "Finder" to set label index of myPath to theMod
					if counter = 0 then --If First item is wrong
						set theIDs to (items 2 through -1 of theIDs)
					else if counter = 1 then --If Second item is wrong
						set theIDs to (item 1 of theIDs) & (items (counter + 1) through -1 of theIDs)
					else if (counter + 1) = (count theIDs) then --If Last item is wrong
						set theIDs to (items 1 through (counter) of theIDs)
					else --Else
						set theIDs to (items 1 through (counter) of theIDs) & (items (counter + 2) through -1 of theIDs)
					end if
					progBar's incrementBy_(4)
				else
					set counter to counter + 1
					progBar's incrementBy_(1)
				end if
			end repeat
			
			theStat's setStringValue_("Entering Picks Into Excel...")
			doEventFetch()
			repeat with c in theIDs
				--doEventFetch()
				--if (my stopper = "stop") then error "STOP"
				log c
				set picks to (do shell script "cat " & quoted form of ((PPath) & (item 1 of c)))
				set numID to (item 2 of c)
				set counter to 0
				repeat with v in theValues
					set counter to counter + 1
					if ((v as number) = (numID as number)) then --Find row number of Number ID
						exit repeat
					end if
				end repeat
				progBar's incrementBy_(1)
				
				set toAdd to {}
				repeat with b from 1 to NumGames --Throw all picks into array toAdd
					--doEventFetch()
					--if (my stopper = "stop") then error "STOP"
					set thecell to ((item b of alphab) & (counter + 2)) as string
					if b < 10 then set end of toAdd to ((characters 4 thru 6 of (paragraph (item b of critLines) of picks)) as string)
					if b > 9 then set end of toAdd to ((characters 5 thru 7 of (paragraph (item b of critLines) of picks)) as string)
				end repeat
				set AppleScript's text item delimiters to ":"
				set WS to (characters 2 thru -1 of (text item 2 of ((paragraph -9 of picks))))
				set LS to (characters 2 thru -1 of (text item 2 of ((paragraph -7 of picks))))
				set AppleScript's text item delimiters to ASTID
				set end of toAdd to WS as string
				set end of toAdd to LS as string
				set BLITZ to (characters 8 thru 10 of ((paragraph -5 of picks)) as string)
				
				repeat with n from 1 to (count toAdd)
					doEventFetch()
					if (my stopper = "stop") then error "STOP"
					set thecell to ((item n of alphab) & (counter + 2)) as string
					tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set value of cell thecell to (item n of toAdd) --Finally update excel sheet!
					if ((item n of toAdd) = BLITZ) then tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set color index of interior object of cell thecell to 28
				end repeat
				progBar's incrementBy_(1)
				set thecell to ((item (NumGames + 7) of alphab) & (counter + 2)) as string
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set value of cell thecell to BLITZ --Don't forget to add blitz!
				progBar's incrementBy_(1)
			end repeat
		on error theError
			if theError = "STOP" then
				set stopNOW to false
			else
				display dialog "EXPORT ERROR: " & theError
				error theError
			end if
		end try
	end EnterData_
	
	on stopNow_(sender)
		log "Stopping..."
		set my stopper to "stop"
	end stopNow_
	
	on CheckWinners_(myInfo)
		set theRads to item 1 of myInfo
		set TheWeek to ((item 2 of myInfo) as string) as number
		set NumGames to item TheWeek of gameNumbers
		set progBar to item 3 of myInfo
		set theStat to item 6 of myInfo
		set WinArray to {}
		
		theStat's setStringValue_("Initiating...")
		doEventFetch()
		tell current application's NSApp to beginSheet_modalForWindow_modalDelegate_didEndSelector_contextInfo_(item 4 of myInfo, item 5 of myInfo, me, "progWinDidEnd:returnCode:contextInfo:", missing value)
		
		try
			progBar's setIndeterminate_(true)
			progBar's startAnimation_(me)
			theStat's setStringValue_("Reading Excel File For Active Cells...")
			doEventFetch()
			repeat with m in (item 1 of myInfo)
				set end of WinArray to ((m's selectedCell()'s title()) as string)
			end repeat
			set activeRows to {}
			set thenum to 3
			repeat --Generate activeRows array of all row numbers with content
				set thecell to "E" & thenum
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set theValue to value of cell thecell
				if theValue = "stop" then exit repeat
				if (theValue = "") = false then set end of activeRows to thenum
				set thenum to thenum + 1
			end repeat
			
			progBar's stopAnimation_(me)
			progBar's setIndeterminate_(false)
			progBar's setMaxValue_(count activeRows)
			progBar's setDoubleValue_(0)
			theStat's setStringValue_("Checking Winners and Calculating Scores...")
			doEventFetch()
			
			repeat with a in activeRows
				set numWrong to 0
				repeat with s from 1 to NumGames
					set thecell to ((item s of alphab) & a) as string
					tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set val to value of cell thecell
					if WinArray does not contain val then
						tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set font color index of font object of cell thecell to 3
						set numWrong to numWrong + 1
					end if
				end repeat
				set thecell to (item (NumGames + 4) of alphab) & a
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set value of cell thecell to (NumGames - numWrong)
				set thecell to (item (NumGames + 7) of alphab) & a
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set val to value of cell thecell
				if WinArray does not contain val then tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set font color index of font object of cell thecell to 3
				progBar's incrementBy_(1)
			end repeat
			tell current application's NSApp to endSheet_returnCode_(item 4 of myInfo, 0)
		on error theError
			display dialog "ERROR: " & theError
			tell current application's NSApp to endSheet_returnCode_(item 4 of myInfo, 0)
			error theError
		end try
	end CheckWinners_
	
	on clearData_(myInfo)
		set TheWeek to ((item 1 of myInfo) as string) as number
		set NumGames to item TheWeek of gameNumbers
		display dialog "Are you sure you would like to continue? Once \"Clear Data\" has started to run, it can not be canceled"
		set progBar to item 2 of myInfo
		set theStat to item 5 of myInfo
		set theNumber to 0
		set theRanges to {}
		
		progBar's setIndeterminate_(true)
		theStat's setStringValue_("Clearing Excel Sheet...")
		tell current application's NSApp to beginSheet_modalForWindow_modalDelegate_didEndSelector_contextInfo_(item 3 of myInfo, item 4 of myInfo, me, "progWinDidEnd:returnCode:contextInfo:", missing value)
		try
			progBar's startAnimation_(me)
			set fromThis to item 1 of alphab
			set toThis to item (NumGames + 2) of alphab
			repeat
				doEventFetch()
				set theNumber to 1 + theNumber
				set D to "D" & theNumber
				tell application "Microsoft Excel" to tell worksheet ("Week (" & TheWeek & ")") of active workbook to set someval to value of cell D
				if someval = "stop" then exit repeat
			end repeat
			set magicNum to (theNumber - 1)
			
			set end of theRanges to fromThis & "3:" & toThis & magicNum
			set end of theRanges to ((item (NumGames + 4) of alphab) & "3:" & (item (NumGames + 4) of alphab) & magicNum)
			set end of theRanges to ((item (NumGames + 7) of alphab) & "3:" & (item (NumGames + 7) of alphab) & magicNum)
			
			tell application "Microsoft Excel"
				tell worksheet ("Week (" & TheWeek & ")") of active workbook
					repeat with l in theRanges
						set color index of interior object of range l to 0
						set font color index of font object of range l to 0
						set value of range l to ""
					end repeat
				end tell
			end tell
			tell current application's NSApp to endSheet_returnCode_(item 3 of myInfo, 0)
			progBar's stopAnimation_(me)
		on error theError
			display dialog "ERROR: " & theError
			tell current application's NSApp to endSheet_returnCode_(item 3 of myInfo, 0)
			progBar's stopAnimation_(me)
			error theError
		end try
	end clearData_
	
	on doEventFetch()
		repeat
			tell current application's NSApp to set theEvent to nextEventMatchingMask_untilDate_inMode_dequeue_(((current application's NSLeftMouseDownMask) as integer) + ((current application's NSKeyDownMask) as integer), missing value, current application's NSEventTrackingRunLoopMode, true)
			if theEvent is missing value then
				exit repeat
			else
				tell current application's NSApp to sendEvent_(theEvent)
			end if
		end repeat
	end doEventFetch
end script