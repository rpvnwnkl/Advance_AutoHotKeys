/*
	Log Class
	By: Ben McClure <ben.mcclur@gmail.com>
*/

;
; Function: Log_initialize
; Description:
;		Creates and opens a log file, initializing it for future operations.
; Syntax: Log_initialize(LogObject, msg = "", overwrite = false)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		msg - (Optional) A message to output to the log when it is initialized
;		overwrite - (Optional) true to overwrite, false to rename existing files
; Return Value:
;		Returns true if initialized successfully, otherwise false
; Example:
;		[code]Log_initialize(Log, "Log file started...")[/code]
;
Log_initialize(LogObject, msg = "", overwrite = false) {
	; Make sure the log is not already initialized
	if (Log_getInitialized(LogObject))
		return false

	; Get the log dir and filename
	dir := Log_dir(LogObject)
	if (!fileName := Log_getFileName(LogObject))
		return false

	; Concatenate to the full log path
	logFile = %dir%\%fileName%

	; Try to rename the file if we shouldn't overwrite
	if ((!overwrite) and FileExist(logFile)) {
		if (!result := Log_move(LogObject))
			return false
	}

	; Create the directory if it doesn't exist
	FileCreateDir, %dir%

	; If we're overwriting, delete the existing file
	FileDelete, %logFile%

	; Create a blank log file
	FileAppend,, %logFile%

	; Set the log to initialized
	Log_setInitialized(LogObject, true)

	; Output initial message if needed
	if (msg)
		Log_output(LogObject, msg)

	return true
}

;
; Function: Log_finalize
; Description:
;		Finalizes a log file, freeing its contents and optionally saving it
; Syntax: Log_finalize(LogObject, msg = "", save = true)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		msg - (Optional) A message to output to the log when it is initialized
; Return Value:
;		Returns true on success, otherwise false
; Example:
;		[code]Log_finalize(Log, "Completed log...")[/code]
;
Log_finalize(LogObject, msg = "", save = true) {
	; Make sure the log is initialized
	if (!Log_getInitialized(LogObject))
		return false

	; Output a final message if needed
	if (msg)
		Log_output(LogObject, msg)

	; Save the log if needed, and fail if the save failed
	if (save) {
		if (!Log_save(LogObject))
			return false
	}

	; Clear the old log Vector
	data := Log_getData(LogObject)
	Vector_clear(data)
	Log_setData(LogObject, data)

	; Set to uninitialized
	Log_setInitialized(LogObject, false)

	return true
}

;
; Function: Log_output
; Description:
;		Outputs a new line to a log object, optionally writing it to file
; Syntax: Log_output(LogObject, msg = "", toFile = false)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		msg - (Optional) A message to output to the log when it is initialized
;		toFile - (Optional) true to immediately write the new line to the log file
; Return Value:
;		true on success, otherwise false
; Example:
;		[code]Log_output(Log, "Example line", true) ; Writes the new line to file[/code]
;
Log_output(LogObject, msg = "", toFile = false) {
	; Make sure the log is initialized
	if (!Log_getInitialized(LogObject))
		return false

	; Get the formatted pre- and post-strings
	preString := Log_formatPreString(LogObject)
	postString := Log_formatPostString(LogObject)

	; Format the message if needed
	msg := Log_replaceStringVars(LogObject, msg)

	; Create the final output string
	output = %preString%%msg%%postString%
	output := String_new(output)

	; Add the output to the data vector
	data := Log_getData(LogObject)
	Vector_add(data, output)
	Log_setData(LogObject, data)

	; Set the file if we're supposed to write
	if Log_getAutoWrite(LogObject) and (!toFile) {
		dir := Log_dir(LogObject)
		if (fileName := Log_getFileName(LogObject))
			toFile := dir . "\" . fileName
	}

	; Output the line to file if needed
	if (toFile) {
		dir := Log_dir(LogObject)
		if (fileName := Log_getFileName(LogObject))
			FileAppend, %output%`r`n, %dir%\%fileName%
	}

	Return true
}

;
; Function: Log_save
; Description:
;		Saves the entire log to file
; Syntax: Log_save(LogObject, fileName = "")
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		fileName - (Optional) If set, saves to the given file instead of
;			the file specified in LogObject
; Return Value:
;		Returns true if saved successfully, otherwise false
; Example:
;		[code]Log_save(Log) ; Saves to file specified in LogObject[/code]
;
Log_save(LogObject, file = "") {
	; Make sure the log is initialized
	if (!Log_getInitialized(LogObject))
		return false

	if (!data := Log_getData(LogObject))
		return false

	; Get the file path from the object if not provided
	if (!file) {
		dir := Log_dir(LogObject)
		if (!fileName := Log_getFileName(LogObject))
			return false
		file = %dir%\%fileName%
	}

	; Clear the file contents
	FileDelete, %file%
	FileAppend,, %file%

	; Save all lines to the file
	Loop,% Vector_size(data)
		logText .= Vector_get(data, A_Index) . "`r`n"

	FileAppend, %logText%, %file%

	return true
}

;
; Function: Log_move
; Description:
;		Moves an existing log file into a numbered series to make room for a new one
; Syntax: Log_move(LogObject, toFile = "")
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		toFile - (Optional) If specified, moves the file here instead
;			of the location specified in the LogObject
; Return Value:
;		Returns true on success, otherwise false
; Example:
;		[code]Log_move(Log)[/code]
;
Log_move(LogObject, toFile = "") {
	; Get the dir and filename
	dir := Log_dir(LogObject)
	if (!fileName := Log_getFileName(LogObject))
		return false

	; Save the full path to the old file
	oldFile = %dir%\%fileName%

	; Get the file to move it to if not specified
	if (!toFile) {
		; Split the file name from its extension
		SplitPath, fileName,,,ext, name

		; get the last position
		loop {
			if (!FileExist(thisFile := dir . "\" . name . "." . A_Index . "." . ext))
				break
		}

		; Set toFile to the first non-existing file found
		toFile := thisFile
	}

	; File should not exist
	if (FileExist(toFile))
		return false

	; Move the file
	FileMove, %oldFile%, %toFile%

	; rotate logs if necessary
	Log_tidy(LogObject)

	return true
}

;
; Function: Log_tidy
; Description:
;		Rotates the backups, removing the oldest backups over %max%
; Syntax: Log_tidy(LogObject, max = "")
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		max - (Optional) The maximum number of files to allow
;			This defaults to the value specified in the LogObject
; Return Value:
;		Returns true if tidy succeeded (even if nothing done), otherwise false
; Remarks:
;		ErrorLevel is set to the number of files deleted (number over max)
; Example:
;		Log_tidy(LogObject, 5) ; Clean up any log file backups over the limit of 5
;
Log_tidy(LogObject, max = "") {
	; Get the max logs
	of (!max)
		max := Log_getTidy(LogObject)

	; Skip the function if max is not set
	if (max = 0 or !max)
		return true

	; Get the dir and filename
	dir := Log_dir(LogObject)
	if (!fileName := Log_getFileName(LogObject))
		return false

	; Split the name and extension
	SplitPath, fileName,,, ext, name

	; Count current backups
	loop {
		if (!FileExist(thisFile := dir . "\" . name . "." . A_Index . "." . ext))
			break
		total := A_Index
	}

	; Don't do anything if there aren't too many backups
	if (total <= max)
		return true

	; How many logs to clean
	toClean := total - max

	; Rotate the logs as necessary
	start := 0
	loop %total% {
		thisFile := dir . "\" . name . "." . A_Index . "." . ext

		if (!FileExist(thisFile))
			continue

		; Delete the item if it's over the limit
		if (A_Index <= toClean) {
			FileDelete, %thisFile%

		; Move all other items down to the beginning
		} else {
			start++
			moveTo := dir . "\" . name . "." . start . "." . ext

			if (FileExist(moveTo))
				continue

			FileMove, %thisFile%, %moveTo%
		}
	}

	return true
}

;
; Function: Log_formatPreString
; Description:
;		Returns a valid preString for Log_output
; Syntax: Log_formatPreString(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the final preString for Log_output
; Remarks:
;		This is an internal function used by Log_output
; Example:
;		[code]Log_formatPreString(Log)[/code]
;
Log_formatPreString(LogObject) {
	; Get the preString template from the object
	if (!preString := Log_getPreString(LogObject))
		return ""

	return Log_replaceStringVars(LogObject, preString)
}

;
; Function: Log_formatPostString
; Description:
;		Returns a valid postString for Log_output
; Syntax: Log_formatPostString(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the final postString for Log_output
; Remarks:
;		This is an internal function used by Log_output
; Example:
;		[code]Log_formatPostString(Log)[/code]
;
Log_formatPostString(LogObject) {
	; Get the preString template from the object
	if (!postString := Log_getPostString(LogObject))
		return ""

	return Log_replaceStringVars(LogObject, postString)
}

;
; Function: Log_replaceStringVars
; Description:
;		Replaces dynamic variables in a log string
; Syntax: Log_replaceStringVars(string)
; Parameters:
;		string - the input string to replace vars inside
; Return Value:
;		Returns the final string for Log_output
; Remarks:
;		This is an internal function used by Log_output
; Example:
;		[code]Log_replaceStringVars(string)[/code]
;
Log_replaceStringVars(LogObject, string) {
	; If no string, return blank
	if (string = "")
		return string

	; Replace time var
	if (InStr(string, "$time")) {
		; Get the saved time template
		timeFormat := Log_getTimeFormat(LogObject)

		; Format the current time
		FormatTime, time, %A_Now%, %timeFormat%

		string := RegExReplace(string, "\$time", time)
	}

	; Replace title var
	if (InStr(string, "$title")) {
		; Get saved title
		if (title := Log_getTitle(LogObject))
			string := RegExReplace(string, "\$title", title)
	}

	return string
}

Log_dir(LogObject) {
	if (!dir := Log_getDir(LogObject))
		dir := A_WorkingDir

	return dir
}

;
; Function: Log_getDir
; Description:
;		Gets the log directory from a log object
; Syntax: Log_getDir(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Example:
;		[code]value := Log_getDir(Log)[/code]
;
Log_getDir(LogObject) {
	return Class_getString(LogObject, "b1")
}

;
; Function: Log_setDir
; Description:
;		Sets the log directory for a log object
; Syntax: Log_setDir(LogObject, dir)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		dir - The (relative or absolute) path to the log directory to set
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setDir(Log, "logs") ; Sets a relative directory[/code]
; Example:
;		[code]Log_setDir(Log, "C:\logs") ; Sets an absolute directory[/code]
;
Log_setDir(LogObject, dir) {
	Class_setString(LogObject, "b1", dir)
}

;
; Function: Log_getFileName
; Description:
;		Gets the filename from a log object
; Syntax: Log_getFileName(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Example:
;		[code]value := Log_getFileName(Log)[/code]
;
Log_getFileName(LogObject) {
	return Class_getString(LogObject, "b2")
}

;
; Function: Log_setFileName
; Description:
;		Sets the filename for a log object
; Syntax: Log_setFileName(LogObject, fileName)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		fileName - The filename (no path) to set for the log object
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setFileName(Log, "log.txt") ; Sets a relative directory[/code]
;
Log_setFileName(LogObject, fileName) {
	Class_setString(LogObject, "b2", fileName)
}

;
; Function: Log_getTitle
; Description:
;		Gets the title from a log object
; Syntax: Log_getTitle(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Example:
;		[code]value := Log_getTitle(Log)[/code]
;
Log_getTitle(LogObject) {
	return Class_getString(LogObject, "b3")
}

;
; Function: Log_setTitle
; Description:
;		Sets the title for a log object
; Syntax: Log_setTitle(LogObject, title)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		title - The title to set for the log object
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setTitle(Log, "My Log File")[/code]
;
Log_setTitle(LogObject, title) {
	Class_setString(LogObject, "b3", title)
}

;
; Function: Log_getTimeFormat
; Description:
;		Gets the AHK-style time format from a log object
; Syntax: Log_getTimeFormat(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Example:
;		[code]value := Log_getTimeFormat(Log)[/code]
;
Log_getTimeFormat(LogObject) {
	return Class_getString(LogObject, "b4")
}

;
; Function: Log_setTimeFormat
; Description:
;		Sets the AHK-style time format for a log object
; Syntax: Log_setTimeFormat(LogObject, timeFormat)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		timeFormat - The AHK-style time format to set for the log object
; Return Value:
;		Returns nothing
; Remarks:
;		This value can be inserted into the PreString or PostString
;		by adding a $t reference in the string
; Example:
;		[code]Log_setTimeFormat(Log, "")[/code]
;
Log_setTimeFormat(LogObject, timeFormat) {
	Class_setString(LogObject, "b4", timeFormat)
}

;
; Function: Log_getPreString
; Description:
;		Gets the preString from a log object
; Syntax: Log_getPreString(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Remarks:
;		The preString is placed at the beginning of every line
;		of the log object, and usually contains at least a
;		date/time reference
; Example:
;		[code]value := Log_getPreString(Log)[/code]
;
Log_getPreString(LogObject) {
	return Class_getString(LogObject, "b5")
}

;
; Function: Log_setPreString
; Description:
;		Sets the string to place before every line in the log object
; Syntax: Log_setPreString(LogObject, preString)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		preString - The string to place before each log line
; Return Value:
;		Returns nothing
; Remarks:
;		Allowed variables:
;		 $time - inserts the time according to the object's timeFormat
;		 $title - inserts the object's title
; Example:
;		[code]Log_setPreString(Log, "[$time] ")[/code]
;
Log_setPreString(LogObject, preString) {
	Class_setString(LogObject, "b5", preString)
}

;
; Function: Log_getPostString
; Description:
;		Gets the postString from a log object
; Syntax: Log_getPostString(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the LogObject, if it exists
; Remarks:
;		The postString is placed at the end of every line
;		of the log object, and is often blank
; Example:
;		[code]value := Log_getPostString(Log)[/code]
;
Log_getPostString(LogObject) {
	return Class_getString(LogObject, "b6")
}

;
; Function: Log_setPostString
; Description:
;		Sets the string to place after every line in the log object
; Syntax: Log_setPostString(LogObject, postString)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		postString - The string to place after each log line
; Return Value:
;		Returns nothing
; Remarks:
;		Allowed variables:
;		 $time - inserts the time according to the object's timeFormat
;		 $title - inserts the object's title
; Example:
;		[code]Log_setPostString(Log, " - $time")[/code]
;
Log_setPostString(LogObject, postString) {
	Class_setString(LogObject, "b6", postString)
}

;
; Function: Log_getData
; Description:
;		Gets the log data set in the log object
; Syntax: Log_getData(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the Vector of log lines from the log object
; Example:
;		[code]value := Log_getData(Log)[/code]
;
Log_getData(LogObject) {
	return Class_getValue(LogObject, "b7", "obj")
}

;
; Function: Log_setData
; Description:
;		Sets the string to place after every line in the log object
; Syntax: Log_setData(LogObject, data)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		data - The vector of log data to store
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setData(Log, logdata)[/code]
;
Log_setData(LogObject, data) {
	Class_setValue(LogObject, "b7", data, "obj")
}

;
; Function: Log_getTidy
; Description:
;		Gets the maximum number of logs set in the log object
; Syntax: Log_getTidy(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the number of logs to keep max (or blank for all)
; Remarks:
;		If this is blank, log backups will not be cleaned automatically
; Example:
;		[code]value := Log_getTidy(Log)[/code]
;
Log_getTidy(LogObject) {
	return Class_getValue(LogObject, "b8", "UChar")
}

;
; Function: Log_setTidy
; Description:
;		Sets the maximum number of log backups for this log object
; Syntax: Log_setTidy(LogObject, tidy)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		tidy - The maximum number of log backups to keep (0 for all)
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setTidy(Log, " - $time")[/code]
;
Log_setTidy(LogObject, tidy) {
	Class_setValue(LogObject, "b8", tidy, "UChar")
}

;
; Function: Log_getAutoWrite
; Description:
;		Gets the autoWrite value for a log object
; Syntax: Log_getAutoWrite(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns the value from the log object
; Example:
;		[code]value := Log_getAutoWrite(Log)[/code]
;
Log_getAutoWrite(LogObject) {
	return Class_getValue(LogObject, "b8a", "UChar")
}

;
; Function: Log_setAutoWrite
; Description:
;		Sets the autoWrite value for a log object
; Syntax: Log_setAutoWrite(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		autoWrite - true to automatically save every line to file,
;			false to wait until the end
; Return Value:
;		Returns nothing
; Example:
;		[code]Log_setAutoWrite(Log, true)[/code]
;
Log_setAutoWrite(LogObject, autoWrite) {
	Class_setValue(LogObject, "b8a", autoWrite, "UChar")
}

;
; Function: Log_getInitialized
; Description:
;		Returns whether or not the log has been initialized
; Syntax: Log_getInitialized(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		Returns true if initialized, otherwise false
; Remarks:
;		Initialized means the file is ready for writing. This is checked automatically
; Example:
;		[code]value := Log_getInitialized(Log)[/code]
;
Log_getInitialized(LogObject) {
	return Class_getValue(LogObject, "b8b", "UChar")
}

;
; Function: Log_setInitialized
; Description:
;		Sets whether the log file has been initialized
; Syntax: Log_setInitialized(LogObject, initialized)
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		initialized - true for initialized, false for uninitialized
; Return Value:
;		Returns nothing
; Remarks:
;		This is usually set internally, and should not be set manually
;
Log_setInitialized(LogObject, initialized) {
	Class_setValue(LogObject, "b8b", initialized, "UChar")
}

;
; Function: Log_setDefaultTimeFormat
; Description:
;		Sets a default time format for all new objects in the class
; Syntax: Log_setDefaultTimeFormat(string = 0)
; Parameters:
;		format - (Optional) The new time format to set, or 0 to just get the value
; Return Value:
;		Returns the current (or new) default time format
; Example:
;		[code]Log_setDefaultTimeFormat("yyyy-MM-dd hh:mm:ss")[/code]
; Example:
;		[code]preString := Log_setDefaultTimeFormat()[/code]
;
Log_setDefaultTimeFormat(format = "") {
	static timeFormat = "yyyy-MM-dd hh:mm:ss tt"

	if (format != "")
		timeFormat := format

	return timeFormat
}

;
; Function: Log_setDefaultPreString
; Description:
;		Sets a default preString template for all new objects in the class
; Syntax: Log_setDefaultPreString(string = 0)
; Parameters:
;		string - (Optional) The new string to set, or 0 to just get the value
; Return Value:
;		Returns the current (or new) default preString template
; Example:
;		[code]Log_setDefaultPreString("[$time] ")[/code]
; Example:
;		[code]preString := Log_setDefaultPreString()[/code]
;
Log_setDefaultPreString(string = 0) {
	static preString = "[$time] "

	if (string != 0)
		preString := string

	return preString
}

;
; Function: Log_setDefaultPostString
; Description:
;		Sets a default postString template for all new objects in the class
; Syntax: Log_setDefaultPostString(string = 0)
; Parameters:
;		string - (Optional) The new string to set, or 0 to just get the value
; Return Value:
;		Returns the current (or new) default postString template
; Example:
;		[code]Log_setDefaultPostString("[$time] ")[/code]
; Example:
;		[code]postString := Log_setDefaultPostString()[/code]
;
Log_setDefaultPostString(string = 0) {
	static postString = ""

	if (string != 0)
		postString := string

	return postString
}

;
; Function: Log_setDefaultTidy
; Description:
;		Sets a default tidy value for new class objects
; Syntax: Log_setDefaultTidy(max = "")
; Parameters:
;		LogObject - The variable referencing a valid Log object
;		max - The new default tidy value to set for new objects
; Return Value:
;		Returns the current (or new) default tidy value
; Example:
;		[code]Log_setDefaultTidy(20)[/code]
; Example:
;		[code]tidy := Log_setDefaultTidy()[/code]
;
Log_setDefaultTidy(max = "") {
	static Tidy = 10

	if (max != "")
		if max is integer
			Tidy := max

	return Tidy
}

;
; Function: Log_new
; Description:
;		Creates a new Log object instance
; Syntax: Log_new(dir = "", fileName = "", title = "", tidy = "", UserDefinedSize = "")
; Parameters:
;		dir - (Optional) The log directory to initially set
;		fileName - (Optional) The log filename to initially set
;		title - (Optional) The log title to initially set
; Return Value:
;		Returns a reference to the new object instance
; Example:
;		Log := Log_new("logs", "log.txt", "My Log File")
;
Log_new(dir = "", fileName = "", title = "") {
	; size of built-in values
	static BuiltInSize = 31
	static UserDefinedSize = 0

	; Use default values if none is specified
	timeFormat := Log_setDefaultTimeFormat()
	preString := Log_setDefaultPreString()
	postString := Log_setDefaultPostString()
	tidy := Log_setDefaultTidy()

	; Allocate Log info structure
	if (LogObject := Class_Alloc(4 * UserDefinedSize + BuiltInSize)) {
		;Store "Class values" (must be present in ALL classes)
		Class_setClass(LogObject, Log_initClass())
		Class_setBuiltInSize(LogObject, BuiltInSize)
		Class_setUserDefinedSize(LogObject, UserDefinedSize)

        ; Set Built-in values
        Log_setDir(LogObject, dir)
        Log_setFileName(LogObject, fileName)
        Log_setTitle(LogObject, title)
        Log_setTimeFormat(LogObject, timeFormat)
        Log_setPreString(LogObject, preString)
        Log_setPostString(LogObject, postString)
        Log_setTidy(LogObject, tidy)
        Log_setData(LogObject, Vector_new())
    }

    return LogObject
}

;
; Function: Log_destroy
; Description:
;		Destroys a created log object (instance of the Log class)
; Syntax: Log_destroy(LogObject)
; Parameters:
;		LogObject - The variable referencing a valid Log object
; Return Value:
;		true if deleted, false if not ready to delete
; Example:
;		Log_destroy(Log)
;
Log_destroy(LogObject) {
	if (!LogObject)
		return

	if !Class_readyToDelete(LogObject)
		return false

    ;Clear built-in fields
    Log_setDir(LogObject, "")
    Log_setFileName(LogObject, "")
    Log_setTitle(LogObject, "")
    Log_setTimeFormat(LogObject, "")
    Log_setPreString(LogObject, "")
    Log_setPostString(LogObject, "")
    Log_setData(LogObject, 0)
    Log_setTidy(LogObject, 0)
    Log_setAutoWrite(LogObject, 0)
    Log_setInitialized(LogObject, 0)

    ;Clear "Class values"
	Class_setBuiltInSize(LogObject, 0)
	Class_setUserDefinedSize(LogObject, 0)

	Class_setClass(LogObject, "")

    ;Free <LogObject>
	Class_Free(LogObject)

	return true
}

;
; Function: Log_initClass
; Description:
;		Initializes the Log class and stores its name
; Syntax: Log_initClass()
; Return Value:
;		Always returns a reference to the class name in memory
; Remarks:
;		This is an internal function that shouldn't need to be called manually
;
Log_initClass()
{
	static ClassName := "Log"

	return &ClassName
}