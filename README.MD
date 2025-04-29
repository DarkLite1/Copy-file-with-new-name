# PowerShell File Copy/Move Script

## Overview

This PowerShell script provides a flexible way to copy or move files from a source directory to a destination directory based on specified criteria. It's designed to be easily configured and automated, making it suitable for scheduled tasks.

The script provides granular control over file selection using:

1.  **Filename Matching:** Employ regular expressions for precise filename filtering.
2.  **Creation Date Filtering:** Process files created within a specified number of recent days.

It then executes the chosen action (copy or move) and includes comprehensive logging and notification capabilities.

## Features

* **Copy or Move Operations:** Choose between copying files (preserving originals) or moving them (deleting originals from the source).
* **Advanced Filename Filtering (Regex):** Utilize regular expressions for sophisticated pattern matching to select target files.
* **Creation Date Filtering:** Process only files created within the last 'N' days, offering flexibility in handling time-sensitive data.
* **Recursive Directory Traversal:** Optionally process files within subdirectories of the source folder.
* **Duplicate File Handling:** Configure whether to overwrite existing files in the destination or log an error upon encountering duplicates.
* **JSON Configuration:** Manage all script parameters through an external JSON configuration file, promoting maintainability and ease of modification.
* **Detailed Logging:** Records script activities, including successes and failures, to specified log files with customizable formats and retention policies.
* **Email Notifications:** Receive email updates based on script execution status (on error, always, or never), with customizable content and SMTP settings.
* **Event Log Integration:** Optionally write script events to the Windows Event Log for centralized monitoring.
* **Concurrency Control:** Manage the number of tasks running in parallel to optimize performance and resource usage.
* **Environment Variable Support:** Leverage environment variables for sensitive information like email credentials and paths.
* **Automation Ready:** Built for seamless integration with task schedulers, requiring appropriate permissions for the executing account.

## Prerequisites

* **Windows PowerShell 7 or later:** Ensure the script is run with a compatible PowerShell version for optimal functionality.
* **Permissions:** The account running the script requires:
    * Read access to the `Source.Folder(s)`.
    * Write/Modify access to the `Destination.Folder`.
    * Write access to the specified `SaveLogFiles.Where.Folder` (if logging is enabled).
    * Permissions to send emails via the configured SMTP server (if email notifications are enabled).
    * Permissions to write to the specified Windows Event Log (if event logging is enabled).
* **(Optional) MailKit and MimeKit Assemblies:** Required for email functionality. The configuration file provides guidance on installing these NuGet packages.
* **(Optional) ImportExcel PowerShell Module:** Required if you intend to use `.xlsx` as a log file extension.

## Configuration

The script is primarily configured using a JSON file, which is passed to the script via a parameter (e.g., `-ConfigurationJsonFile "config.json"`). This approach keeps the script logic separate from the configuration, making it easier to manage different scenarios.

### JSON Configuration File (`-ConfigurationJsonFile`)

Create a `.json` file with the following structure. Refer to the `?` section within the `Example.json` file for detailed explanations of each option.

```json
{
  "Tasks" : [
    {
      "Action": "copy", // or "move"
      "ProcessFilesCreatedInTheLastNumberOfDays": 7, // 0 - all files; 1 - files created today; 5 - created today and in the last 4 days; ...
      "Source": {
        "Folder": "C:\\folderA",
        "Recurse": true, // true - include subfolders; false - only the root folder
        "MatchFileNameRegex": ".*\\.csv$" // Regular expression for filename matching. '.*' - all files; '.*\\.csv$' - csv files; ...
      },
      "Destination": {
        "Folder": "D:\\folderB",
        "OverWriteFile": true // true - overwrite existing files; false - log error if file exists
      },
    }
  ]
}