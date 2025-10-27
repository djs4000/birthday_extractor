# AGENTS.md

This document provides a high-level overview of the Birthday Extractor project to give context to AI agents and new developers.

## Project Summary

**Birthday Extractor** is a C# Windows Forms (.NET) application designed to process customer data to identify children with upcoming birthdays. It serves as a tool for marketing and customer relationship management (CRM) by generating targeted lead lists.

The core functionality includes:
1.  **Data Ingestion**: It can read customer data from two sources:
    *   A local CSV file.
    *   A remote online API endpoint.
2.  **Filtering**: It filters the customer list to find children whose birthdays fall within a user-specified date range and whose age on their birthday falls within a configured minimum and maximum age (e.g., 3-14 years old).
3.  **Data Processing**: It performs data normalization, particularly for phone numbers (using `libphonenumber-csharp`), and attempts to link child records to their parent/guardian records from the same dataset based on shared email or phone numbers.
4.  **Output Generation**: The filtered list of birthday leads can be exported as:
    *   A CSV file.
    *   An XLSX (Excel) file with table formatting.
5.  **ERP Integration**: It has an optional feature to upload the generated leads directly into an ERPNext instance as "Lead" documents, checking for duplicates to avoid re-inserting existing leads.
6.  **Silent/CLI Mode**: The application can be run from the command line for automated, headless operation.
7.  **Auto-Update**: It includes a mechanism to check for new versions on GitHub and perform an in-place update.

## Key Components & Architecture

The project is structured into several key classes, each with a distinct responsibility.

### Application Entry & UI

*   `Program.cs`: The main entry point. It handles parsing command-line arguments for silent mode and initializes the Windows Forms application if no CLI arguments are provided.
*   `MainForm.cs`: The primary UI window. It's built programmatically (without a `.designer.cs` file). It allows the user to configure and trigger the extraction process, view logs, access settings, and see processing history. It orchestrates calls to the processing and upload logic.

### Core Logic

*   `Processing.cs`: This is the heart of the application. The `Processing` class contains the entire data processing pipeline. It is stateless and takes a `ProcOptions` object as input. Its `Process` method performs the following steps:
    1.  Reads data from the specified source (CSV or online).
    2.  Normalizes raw data into an internal `Row` representation.
    3.  Parses dates of birth and normalizes phone numbers.
    4.  Identifies adults (potential guardians) and children.
    5.  Filters for children with birthdays in the target window and age range.
    6.  Correlates children with their guardians.
    7.  Generates a unique `BusinessKey` for each lead to assist with deduplication.
    8.  Writes the final, sorted list to CSV and/or XLSX files.
    9.  Returns a `ProcResult` object containing the outcome and the list of `ExtractedLead` objects.

### Configuration & State

*   `AppConfig.cs`: Defines the `AppConfig` data model, which stores all user preferences, API credentials, and run history.
*   `ConfigStore.cs`: A static helper class responsible for serializing `AppConfig` to and from a `config.json` file located in the user's `%LOCALAPPDATA%\BirthdayExtractor` directory. It also includes a utility for computing file hashes.

### Integrations

*   `ErpNextClient.cs`: A low-level client for interacting with the ERPNext REST API. It handles authentication and the specifics of fetching existing leads (to check for duplicates) and creating new ones.
*   `ErpNextUploader.cs`: A high-level orchestrator that uses `ErpNextClient` to manage the batch upload process. It filters out leads with missing required data, checks for existing leads using the `BusinessKey`, and reports a summary of the operation.
*   `UpdateService.cs` (referenced, not provided): A class responsible for querying the GitHub API to check for new application releases.
*   `SelfUpdateCoordinator.cs`: Manages the tricky process of the application replacing its own executable file after an update is downloaded.

### Utilities

*   `LogRouter.cs`: A static class providing a centralized logging sink. It can queue log messages and forward them to the `MainForm`'s log view once it's available, allowing non-UI components to provide user-visible feedback.
*   `AppVersion.cs`: A simple static class that centralizes the application's version string (`Display`) and its parsed `System.Version` equivalent (`Semantic`).

## Key Data Models

*   `ProcOptions`: An object that encapsulates all parameters for a single processing run (e.g., dates, paths, ages, output flags).
*   `ProcResult`: An object that summarizes the results of a processing run, including counts and paths to output files.
*   `ExtractedLead`: A clean, flattened data transfer object (DTO) representing a single child/guardian lead, used for the ERPNext upload.
*   `ProcessedWindow`: A record stored in `AppConfig.History` that logs a completed run (date range, source file info, and result count) to help users avoid duplicate work.

## Primary Workflows

1.  **GUI Workflow**: The user interacts with `MainForm` to select a CSV or use the configured online source, sets the date and age parameters, and clicks "Run". `MainForm` calls `Processing.Process`. If the run is successful and produces leads, the "Upload to ERPNext" button is enabled.

2.  **CLI Workflow**: The user runs the executable with command-line flags like `--silent`, `--csv <path>`, `--start <date>`, and `--end <date>`. `Program.cs` parses these arguments, builds a `ProcOptions` object, and directly calls `Processing.Process`. It can also trigger an ERPNext upload if requested via flags.

3.  **Update Workflow**: On startup, `MainForm` checks for updates via the `UpdateService`. If a new version is found and the user agrees, it's downloaded. A batch script is created to replace the old executable with the new one and relaunch the application. `SelfUpdateCoordinator` handles the logic for the newly launched executable to put itself in the correct installation directory.