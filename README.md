# Pearascope V2
This app provides a framework for performing analysis on a WPILOG file, in our case as output from our use of AdvantageKit.  The expected approach is to iterate through log entries during teleop (and possibly auto), seeking entries of interest like specific button presses or sensor values, to tell a story of what happened throughout a match.  The analysis directly loads the WPILOG file using WPILib DataLogReader, and output goes directly to an Excel file.

## Year-over-year Updates
Each year should have a PearascopeYYYY processor class where the unique analysis for that year will be performed.  All other files and classes are for orchestration and utility/helper functions, which should not need to change each year.

## Explanation of Each Class
### Main.java
Houses the main() function for the app.  Processes the command line arguments and dispatches to the appropriate processor class.

#### Command Line Arguments
|Argument|Purpose|
|---|---|
|-year YYYY|Can be provided to specify which processor to use.  If not provided, app defaults to the current year.|
|-raw|For each log file, also generate a raw dump of all log entries during auto and teleop periods (cannot be used with -monitor)|
|-monitor|Constantly monitor for new files in the ./input folder (this is the primary intended use case) |
|&lt;file paths&gt;|One or more file paths separated by a space - paths can be relative or absolute.  If no paths are provided, all files in ./input will be processed.|

### PearascopeProcessor.java
Abstract base class for all processors.  Defines various abstract functions that must be implemented by processor classes, and directly implements the following:
* _run()_: Orchestrates whether to monitor the input folder for changes, process all files in the input folder (then quit), or process file paths provided in the command line.
* _processLogs()_: Iterates through the provided file paths to process each log one at a time.
* _processLog()_: Opens the log file in a WPILib DataLogReader to pass to the derived processor class.

### PearascopeYYYY.java
All named functions below are defined in the PearascopeProcessor base class as abstract functions that must be implemented in this derived class.
* _processLog()_: Sets up the output Excel file, iterates through the log entries, writes entries of interest to the output file, adds various analysis, and applies custom formatting to the sheet.
* _outputEntriesOfInterest()_: Determines whether the name of the entry is of interest and adds a row to the sheet with appropriate info from the entry.
* _addOutputRow()_: Uses Excel helper functions to add the row (if we ever standardize analysis format, this can probably be moved to Excel.java).
* _Analysis functions (e.g. addIntakeTimes())_: Each such function iterates through the output rows one by one to add to the story.  Most of these will find a start and end point of some action and add a formula to display the time spent for that action.
* _formatOutput()_: Applies custom formatting and filtering to the output sheet.
* _dumpRawLog()_: Creates a separate "RAW" output Excel file with all entries from the log during auto and teleop periods.

### ColumnEnum.java
Unless and until we standardize analysis output format, a per-year COLUMN enum is needed.  This interface allows us to pass enums implementing this interface to helper functions.  If we standardize analysis, we should be able to move the enum to Excel.java and eliminate this interface.

Each PearascopeYYYY class will need an enum defined to represent each column in the spreadsheet, listed sequentially as they should appear in the spreadsheet left-to-right.

### Excel.java
Helper class for using Apache POI to read and write Excel files.

### Utils.java
Miscellaneous helper functions.

### Known Issues
* An issue occurs when trying to run/debug the app from VS Code where the second run without recompiling triggers an error.  This is partially due to a bug in the Gradle extension, necessitating a switch to the pre-prelease version of that extension (which fixes the bug).  Although that extension should eventually release the fix, the error still occurs as there is some implicit tie to Java 21.  But, we need Java 17 due to WPILib being compiled against version 17.  The fix appears to be to include an explicit reference to Java 17 in .vscode/settings.json, but this requires the path to your Java 17 home - yours may be different from what is here, so adjust as needed (and I suggest you do not commit changes to this file).