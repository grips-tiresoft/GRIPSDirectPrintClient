-- GRIPS Direct Print Launcher AppleScript
-- This handles opening .grdp files via macOS document events

on run
	display dialog "No file provided to GRIPS Direct Print." & return & return & "Please open a .grdp file." buttons {"OK"} default button "OK" with icon stop with title "GRIPS Direct Print"
end run

on open theFiles
	try
		-- Show that we received the file
		display notification "Processing " & (count of theFiles) & " file(s)" with title "GRIPS Direct Print"
		
		-- Get the path to this application
		set appPath to path to me as text
		set appPosixPath to POSIX path of appPath
		
		-- Remove trailing slash and /Contents/MacOS/applet from path
		set resourcesPath to appPosixPath & "Contents/Resources/"
		set printScript to resourcesPath & "Print-GRDPFile.sh"
		
		-- Use user's Library/Logs for debug logging
		set logFile to (POSIX path of (path to library folder from user domain)) & "Logs/GRIPSDirectPrint.log"
		
	-- Create log directory
	do shell script "mkdir -p " & quoted form of ((POSIX path of (path to library folder from user domain)) & "Logs")
	
	-- Rotate log if needed (10MB limit, keep 5 rotations)
	do shell script "
if [ -f " & quoted form of logFile & " ]; then
    filesize=$(stat -f%z " & quoted form of logFile & " 2>/dev/null || echo 0)
    sizemb=$((filesize / 1024 / 1024))
    if [ $sizemb -ge 5 ]; then
        [ -f " & quoted form of logFile & ".5 ] && rm -f " & quoted form of logFile & ".5
        for i in 4 3 2 1; do
            [ -f " & quoted form of logFile & ".$i ] && mv " & quoted form of logFile & ".$i " & quoted form of logFile & ".$((i+1))
        done
        mv " & quoted form of logFile & " " & quoted form of logFile & ".1
        touch " & quoted form of logFile & "
    fi
fi
"
	
	-- Write debug info to log (append mode)
	do shell script "echo '=== Run started at '$(date)' ===' >> " & quoted form of logFile
	do shell script "echo 'Resources Path: " & resourcesPath & "' >> " & quoted form of logFile
		-- Process each dropped file
		repeat with aFile in theFiles
			set filePath to POSIX path of aFile
			
			do shell script "echo 'Processing: " & filePath & "' >> " & quoted form of logFile
			
			-- Execute the print script with the file
			try
				-- Set environment variable to prevent script from calling 'open'
				-- Add common paths where jq might be installed
				set shellCommand to "export PATH=/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH && export GRDP_NO_OPEN=1 && cd " & quoted form of resourcesPath & " && " & quoted form of printScript & " -i " & quoted form of filePath & " 2>&1"
				
				do shell script "echo 'About to execute shell command' >> " & quoted form of logFile
				set output to do shell script shellCommand
				do shell script "echo 'Shell command completed' >> " & quoted form of logFile
				do shell script "echo 'Output length: " & (length of output) & "' >> " & quoted form of logFile
				
				-- Write output to log file
				do shell script "echo 'Script Output:' >> " & quoted form of logFile
				do shell script "echo " & quoted form of output & " >> " & quoted form of logFile
				
				-- Parse output for files that need to be opened
				-- Look for lines like "Opening file: /path/to/file.eml"
				set outputLines to paragraphs of output
				repeat with aLine in outputLines
					if aLine contains "Opening file: " then
						set openFilePath to text ((offset of "Opening file: " in aLine) + 14) thru -1 of aLine
						try
							-- Use AppleScript to open the file
							do shell script "open " & quoted form of openFilePath
							display notification "Opened: " & openFilePath with title "GRIPS Direct Print"
						on error openErr
							display dialog "Failed to open file:" & return & openFilePath & return & return & "Error: " & openErr buttons {"OK"} with icon caution
						end try
					end if
				end repeat
				
				-- Show success notification
				display notification "File processed successfully" with title "GRIPS Direct Print"
			on error errMsg
				display dialog "Error processing file:" & return & return & errMsg & return & return & "Script: " & printScript & return & "File: " & filePath buttons {"OK"} default button "OK" with icon stop with title "GRIPS Direct Print Error"
			end try
		end repeat
	on error errMsg
		display dialog "Error in launcher:" & return & return & errMsg buttons {"OK"} default button "OK" with icon stop with title "GRIPS Direct Print Error"
	end try
end open
