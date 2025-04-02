# PowerShell File Copy/Move Script

## Overview

This PowerShell script provides a flexible way to copy or move files from a source directory to a destination directory based on specified criteria. It's designed to be easily configured and automated, making it suitable for scheduled tasks.

The script selects files based on:
1.  A filename pattern (using regular expressions or wildcards).
2.  The file's creation date (processing files created within a specific number of recent days).

It then performs either a 'copy' or 'move' operation. Error logging is implemented to capture any issues during execution.

## Features

* **Copy or Move:** Choose whether to copy files (leaving originals) or move them (removing originals).
* **Filename Filtering:** Use regex or wildcard patterns (`*.txt`, `data_*.csv`, etc.) to select specific files.
* **Date Filtering:** Process only files created within the last 'N' days.
* **Recursive Search:** Optionally include files from subdirectories within the source folder.
* **Overwrite Control:** Decide whether to overwrite files with the same name in the destination folder or log an error.
* **Configuration via JSON:** Manage all script parameters easily through a separate JSON configuration file.
* **Error Logging:** Logs only errors encountered during the process to a specified log folder.
* **Automation Ready:** Intended for use with scheduled tasks, requiring appropriate folder permissions for the task's execution account.

## Prerequisites

* Windows PowerShell 7.
* Permissions:
    * Read access to the `Source.Folder`.
    * Write/Modify access to the `Destination.Folder`.
    * Write access to the `LogFolder`.

## Configuration

The script primarily uses a JSON file for configuration, passed via the `-ImportFile` parameter. This allows for easy management of settings without modifying the script itself.

### JSON Configuration File (`-ImportFile`)

Create a `.json` file with the following structure:

```json
{
  "Action": "copy", // or "move"
  "Source": {
    "Folder": "C:\\Path\\To\\Source",
    "Recurse": true, // or false
    "MatchFileNameRegex": "*.csv" // e.g., "*.txt", "report_*.docx", "*.*"
  },
  "Destination": {
    "Folder": "D:\\Path\\To\\Destination",
    "OverWriteFile": false // or true
  },
  "ProcessFilesCreatedInTheLastNumberOfDays": 7 // 0 = process all, >0 = process files created in the last N days
}