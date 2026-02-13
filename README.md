# Schedules Excel Export - Revit Add-in

A Revit add-in that enables users to export multiple schedules from Autodesk Revit to Excel (.xlsx) format with advanced formatting and filtering options.

## Features

- **Batch Export**: Export multiple Revit schedules to a single Excel workbook
- **Flexible File Handling**: Create new Excel files or append to existing ones
- **Smart Formatting Options**:
  - Write numbers as strings or preserve numeric types
  - Exclude empty rows automatically
  - Filter out revision schedules by default
- **User-Friendly Interface**: 
  - Select specific schedules to export via checklist
  - "Select All" option for quick selection
  - Real-time file path validation
- **Automatic File Management**: 
  - Opens the exported file automatically after export
  - Detects and warns if target file is already open
  - Creates directories if they don't exist

## Requirements

- Autodesk Revit (2015 or later recommended)
- .NET Framework compatible with your Revit version

## Installation

### Option 1: Using the Installer (Recommended)

1. Download the latest release from the [Releases](https://github.com/ancodia/schedules-excel-export-revit-addin/releases) page
2. Run the InnoSetup installer
3. Restart Revit
4. The add-in will appear in the External Tools menu

### Option 2: Manual Installation

1. Clone or download this repository
2. Build the solution in Visual Studio
3. Copy the compiled DLL and `.addin` manifest file to:
   ```
   C:\ProgramData\Autodesk\Revit\Addins\[YEAR]\
   ```
4. Restart Revit

## Usage

### Basic Workflow

1. **Launch the Add-in**
   - In Revit, go to Add-Ins tab → External Tools
   - Click "Schedule Exporter" (or the command name defined in your `.addin` file)

2. **Choose Export Destination**
   - Click **"New File"** to create a new Excel file, or
   - Click **"Select File"** to append to an existing Excel file

3. **Select Schedules**
   - Check the schedules you want to export from the list
   - Use "Select All" checkbox for quick selection
   - Revision schedules are automatically filtered out

4. **Configure Export Options**
   - **Write numbers as strings**: 
     - ✅ Checked (default): All cell values written as text
     - ❌ Unchecked: Numeric values preserved as numbers in Excel
   - **Exclude all empty rows**: 
     - ✅ Checked: Removes all rows with no data
     - ❌ Unchecked (default): Only removes empty header rows

5. **Export**
   - Click **"Write"** to export
   - The Excel file will open automatically

### Export Behavior

#### Schedule Organization
- Each schedule is exported to a separate worksheet in the Excel workbook
- Worksheet names match the Revit schedule names (sanitized for Excel compatibility)
- Existing worksheets with matching names are overwritten

#### Data Processing
- **Header Rows**: The first row is always the header row from the schedule
- **Empty Row Handling**: 
  - Empty header rows (row index 1) are always excluded
  - Additional empty rows can be excluded if "Exclude all empty rows" is enabled
- **Cell Values**: Trimmed of whitespace automatically

#### Worksheet Naming
Invalid Excel characters in schedule names are automatically removed:
- `: / \ [ ] * ?`

For example:
- `"Schedule: Doors"` → `"Schedule Doors"`
- `"Area/Room Schedule"` → `"AreaRoom Schedule"`

## Code Structure

### Main Components

#### `Command.cs` - Core Logic
- **Entry Point**: `Execute()` - Main command execution triggered by Revit
- **Schedule Collection**: Filters and sorts all schedules in the active document
- **Export Engine**: `ExportToExcel()` - Handles Excel file creation/modification
- **Data Preparation**: `PrepareScheduleData()` - Converts Revit schedule data to Excel-compatible format
- **Utility Methods**:
  - `SanitizeWorksheetName()` - Removes invalid Excel worksheet characters
  - `IsRowEmpty()` - Detects empty rows for optional filtering

#### `ExportForm.cs` - User Interface
- Windows Forms-based dialog
- File selection (new/existing)
- Schedule selection with checklist
- Export configuration options
- Input validation and error handling

## Technical Details

### Schedule Filtering

The add-in automatically excludes revision schedules from the export list:

```csharp
.Where(schedule => !schedule.Name.Contains("<Revision"))
```

### Numeric Type Handling

When "Write numbers as strings" is unchecked, the add-in attempts to preserve numeric types:

```csharp
if (double.TryParse(cellValue, out double doubleValue))
    rowData.Add(doubleValue);
else if (int.TryParse(cellValue, out int intValue))
    rowData.Add(intValue);
else
    rowData.Add(cellValue);
```

### File Safety

The add-in includes file access validation:
- Checks if target file is currently open
- Creates parent directories if they don't exist
- Loads existing files without overwriting other worksheets

## Known Limitations

- Revision schedules are excluded from export
- Schedule formatting (colors, fonts, merged cells) is not preserved
- Sheet names are limited to Excel's 31-character limit (truncated if necessary)
- Schedule images/graphics are not exported

## Development

### Building from Source

1. Clone the repository:
   ```bash
   git clone https://github.com/ancodia/schedules-excel-export-revit-addin.git
   ```

2. Open `SchedulesExcelExport.sln` in Visual Studio

3. Restore NuGet packages:
   ```bash
   dotnet restore
   ```

4. Build the solution:
   ```bash
   dotnet build
   ```

### Project Structure

```
schedules-excel-export-revit-addin/
├── src/
│   ├── Command.cs              # Main command logic
│   ├── ScheduleExporter.cs     # UI form
│   ├── Helper.cs               # Utility functions
│   └── *.csproj                # Project file
├── InnoSetup/                  # Installer configuration
├── .github/workflows/          # CI/CD workflows
└── SchedulesExcelExport.sln    # Solution file
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Support

For issues, questions, or suggestions, please open an issue on the [GitHub repository](https://github.com/ancodia/schedules-excel-export-revit-addin/issues).

## Changelog

### v1.0.1 (Latest)
- Latest stable release

See the [Releases](https://github.com/ancodia/schedules-excel-export-revit-addin/releases) page for full version history.

## Acknowledgments

- Built using the Autodesk Revit API
- Inspired by the need for efficient schedule data management workflows

---

**Note**: This add-in is not affiliated with or endorsed by Autodesk, Inc.
