using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Resources;
using System.Xml;
using System.Xml.Linq;
using Color = System.Drawing.Color;

namespace Excel2Resx;

public partial class MainForm : Form
{
    private string excelFilePath = string.Empty;
    private string resxFolderPath = string.Empty;
    private string resxFileNamePrefix = "Resource";
    private bool createBackup = true;
    private Stack<UndoAction> undoStack = new Stack<UndoAction>();

    // Registry keys for storing settings
    private const string RegistryKeyPath = @"SOFTWARE\Excel2Resx";
    private const string ResxFolderPathKey = "ResxFolderPath";
    private const string ResxFileNamePrefixKey = "ResxFileNamePrefix";
    private const string CreateBackupKey = "CreateBackup";

    private class UndoAction
    {
        public string ActionType { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string BackupPath { get; set; } = string.Empty;

        public UndoAction(string actionType, string filePath, string backupPath = "")
        {
            ActionType = actionType;
            FilePath = filePath;
            BackupPath = backupPath;
        }
    }

    public MainForm()
    {
        InitializeComponent();

        // Enable drag and drop functionality
        this.AllowDrop = true;
        txtExcelPath.AllowDrop = true;

        // Register event handlers for drag and drop
        this.DragEnter += MainForm_DragEnter;
        this.DragDrop += MainForm_DragDrop;
        txtExcelPath.DragEnter += MainForm_DragEnter;
        txtExcelPath.DragDrop += MainForm_DragDrop;

        // Load settings from registry
        LoadSettings();

        // Initialize the UI state
        UpdateProcessButton();
        UpdateStatusMessage("Ready");

        // Display welcome message
        AppendToLog("Application started. Please select an Excel file with translations.");
    }

    private void LoadSettings()
    {
        try
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (key != null)
                {
                    // Load RESX folder path
                    resxFolderPath = (key.GetValue(ResxFolderPathKey) as string) ?? string.Empty;
                    txtResxFolderPath.Text = resxFolderPath;

                    // Load RESX file name prefix
                    resxFileNamePrefix = (key.GetValue(ResxFileNamePrefixKey) as string) ?? "Resource";
                    txtResxFilePrefix.Text = resxFileNamePrefix;

                    // Load checkbox states
                    createBackup = (key.GetValue(CreateBackupKey) as int?) == 1;
                    chkCreateBackup.Checked = createBackup;
                }
            }
        }
        catch (Exception ex)
        {
            // Silently ignore any errors when loading settings
            AppendToLog($"Warning: Could not load settings - {ex.Message}");
        }
    }

    private void SaveSettings()
    {
        try
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                if (key != null)
                {
                    // Save RESX folder path
                    key.SetValue(ResxFolderPathKey, resxFolderPath);

                    // Save RESX file name prefix
                    key.SetValue(ResxFileNamePrefixKey, resxFileNamePrefix);

                    // Save checkbox states
                    key.SetValue(CreateBackupKey, createBackup ? 1 : 0);
                }
            }
        }
        catch (Exception ex)
        {
            // Silently ignore any errors when saving settings
            AppendToLog($"Warning: Could not save settings - {ex.Message}");
        }
    }

    private void BtnBrowseExcel_Click(object sender, EventArgs e)
    {
        using var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel Files|*.xlsx",
            Title = "Select an Excel file"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            excelFilePath = openFileDialog.FileName;
            txtExcelPath.Text = excelFilePath;
            UpdateProcessButton();

            AppendToLog($"Selected Excel file: {Path.GetFileName(excelFilePath)}");
            UpdateStatusMessage("Ready - Excel file selected");
        }
    }

    private void BtnBrowseResxFolder_Click(object sender, EventArgs e)
    {
        using var folderBrowserDialog = new FolderBrowserDialog
        {
            Description = "Select RESX files folder"
        };

        // Set initial directory if we have one saved
        if (!string.IsNullOrEmpty(resxFolderPath) && Directory.Exists(resxFolderPath))
        {
            folderBrowserDialog.InitialDirectory = resxFolderPath;
        }

        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            resxFolderPath = folderBrowserDialog.SelectedPath;
            txtResxFolderPath.Text = resxFolderPath;
            UpdateProcessButton();

            // Scan for existing RESX files to auto-detect the file prefix
            ScanForExistingResxFiles();

            AppendToLog($"Selected RESX folder: {resxFolderPath}");
            UpdateStatusMessage("Ready - RESX folder selected");

            // Save settings when folder path changes
            SaveSettings();
        }
    }

    private void UpdateProcessButton()
    {
        btnProcess.Enabled = !string.IsNullOrEmpty(excelFilePath) && !string.IsNullOrEmpty(resxFolderPath);

        // Change appearance based on enabled state
        if (btnProcess.Enabled)
        {
            btnProcess.BackColor = Color.FromArgb(0, 120, 215); // Blue when enabled
            btnProcess.ForeColor = Color.White;
        }
        else
        {
            btnProcess.BackColor = Color.FromArgb(230, 230, 230); // Light gray when disabled
            btnProcess.ForeColor = Color.Gray;
        }
    }

    private async void BtnProcess_Click(object sender, EventArgs e)
    {
        try
        {
            // Update UI for processing state
            SetProcessingState(true, "Processing translations...");

            // Clear the undo stack before starting a new operation
            undoStack.Clear();
            btnUndo.Enabled = false;

            // Save current settings before processing
            SaveSettings();

            // Process the Excel file and update/create RESX files
            await Task.Run(() => ProcessExcelToResx());

            // Enable the undo button if there are actions to undo
            btnUndo.Enabled = undoStack.Count > 0;

            // Show completion message
            MessageBox.Show("Translation processing completed successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            UpdateStatusMessage("Ready - Processing completed");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            AppendToLog($"ERROR: {ex.Message}");
            UpdateStatusMessage("Error during processing");
        }
        finally
        {
            // Restore UI after processing
            SetProcessingState(false);
        }
    }

    private void ChkCreateBackup_CheckedChanged(object sender, EventArgs e)
    {
        createBackup = chkCreateBackup.Checked;
        SaveSettings();
    }

    private void TxtResxFilePrefix_TextChanged(object sender, EventArgs e)
    {
        resxFileNamePrefix = txtResxFilePrefix.Text.Trim();

        // If the prefix is empty, set it back to the default
        if (string.IsNullOrWhiteSpace(resxFileNamePrefix))
        {
            resxFileNamePrefix = "Resource";
            txtResxFilePrefix.Text = resxFileNamePrefix;
        }

        SaveSettings();
    }

    private async void BtnUndo_Click(object sender, EventArgs e)
    {
        if (undoStack.Count == 0)
        {
            MessageBox.Show("Nothing to undo.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            // Update UI for processing state
            SetProcessingState(true, "Undoing changes...");

            // Process all undo actions at once
            int undoCount = await Task.Run(() =>
            {
                int count = 0;
                while (undoStack.Count > 0)
                {
                    UndoAction action = undoStack.Pop();

                    switch (action.ActionType)
                    {
                        case "Create":
                            // Delete the created file
                            if (File.Exists(action.FilePath))
                            {
                                File.Delete(action.FilePath);
                                AppendToLog($"Undone: Deleted created file {Path.GetFileName(action.FilePath)}");
                                count++;
                            }
                            break;

                        case "Modify":
                            // Restore from backup
                            if (File.Exists(action.BackupPath))
                            {
                                File.Copy(action.BackupPath, action.FilePath, true);
                                AppendToLog($"Undone: Restored {Path.GetFileName(action.FilePath)} from backup");
                                count++;
                            }
                            break;
                    }
                }
                return count;
            });

            AppendToLog($"Undo completed: {undoCount} file(s) processed");
            UpdateStatusMessage("Ready - Changes undone");

            // Update button state
            btnUndo.Enabled = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error during undo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            AppendToLog($"ERROR during undo: {ex.Message}");
            UpdateStatusMessage("Error during undo operation");
        }
        finally
        {
            // Restore UI after processing
            SetProcessingState(false);
        }
    }

    private void MainForm_DragEnter(object? sender, DragEventArgs e)
    {
        // Check if the dragged data is a file
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data is string[] files)
            {
                // Only accept Excel files
                if (files.Length == 1 && Path.GetExtension(files[0]).ToLowerInvariant() == ".xlsx")
                {
                    e.Effect = DragDropEffects.Copy;  // Show copy icon

                    // Change TextBox appearance during drag
                    if (sender == txtExcelPath)
                    {
                        txtExcelPath.BackColor = Color.FromArgb(230, 240, 255); // Light blue during drag
                    }

                    return;
                }
            }
        }

        e.Effect = DragDropEffects.None;  // Reject the drop
    }

    private void MainForm_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data is string[] files)
            {
                if (files.Length == 1 && Path.GetExtension(files[0]).ToLowerInvariant() == ".xlsx")
                {
                    excelFilePath = files[0];
                    txtExcelPath.Text = excelFilePath;
                    UpdateProcessButton();

                    AppendToLog($"Excel file loaded: {Path.GetFileName(excelFilePath)}");
                    UpdateStatusMessage("Ready - Excel file loaded");
                }
            }
        }

        // Reset TextBox appearance
        txtExcelPath.BackColor = Color.WhiteSmoke;
    }

    private void SetProcessingState(bool isProcessing, string message = "")
    {
        if (InvokeRequired)
        {
            Invoke(new Action<bool, string>(SetProcessingState), isProcessing, message);
            return;
        }

        if (isProcessing)
        {
            // Disable controls during processing
            btnProcess.Enabled = false;
            btnBrowseExcel.Enabled = false;
            btnBrowseResxFolder.Enabled = false;
            txtResxFilePrefix.Enabled = false;
            chkCreateBackup.Enabled = false;

            // Update buttons appearance
            btnProcess.BackColor = Color.FromArgb(200, 200, 200);
            btnProcess.ForeColor = Color.Gray;

            // Set status message
            if (!string.IsNullOrEmpty(message))
            {
                UpdateStatusMessage(message);
            }

            // Change cursor for the whole form
            this.Cursor = Cursors.WaitCursor;
        }
        else
        {
            // Re-enable controls
            btnBrowseExcel.Enabled = true;
            btnBrowseResxFolder.Enabled = true;
            txtResxFilePrefix.Enabled = true;
            chkCreateBackup.Enabled = true;

            // Update process button state based on inputs
            UpdateProcessButton();

            // Reset cursor
            this.Cursor = Cursors.Default;
        }
    }

    private void UpdateStatusMessage(string message)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<string>(UpdateStatusMessage), message);
            return;
        }

        statusLabel.Text = message;
    }

    private void AppendToLog(string message)
    {
        // Invoke on UI thread if needed
        if (InvokeRequired)
        {
            Invoke(new Action<string>(AppendToLog), message);
            return;
        }

        txtLog.AppendText($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
        txtLog.ScrollToCaret();
    }

    private void ProcessExcelToResx()
    {
        AppendToLog("Starting translation processing...");
        UpdateStatusMessage("Reading Excel file...");

        // Dictionary to store the translations for each language
        // Key: Language code (column header), Value: Dictionary of resource key-value pairs
        var translations = new Dictionary<string, Dictionary<string, string>>();

        AppendToLog($"Opening Excel file: {Path.GetFileName(excelFilePath)}");

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Get shared string table for looking up string values
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart!;
            SharedStringTable sharedStringTable = stringTablePart.SharedStringTable;

            // Get all rows
            var rows = sheetData.Elements<Row>().ToList();
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("No data found in the Excel file");
            }

            AppendToLog($"Found {rows.Count} rows in the Excel file");

            // Process header row to get language codes
            var headerRow = rows[0];
            var headerCells = headerRow.Elements<Cell>().ToList();

            // Dictionary to map column index to language code
            var columnLanguageMap = new Dictionary<string, string>();

            // Start from the second cell (index 1) - first is ResourceKey
            for (int i = 1; i < headerCells.Count; i++)
            {
                var cell = headerCells[i];
                string languageCode = GetCellValue(cell, sharedStringTable).Trim();

                if (!string.IsNullOrEmpty(languageCode))
                {
                    string columnReference = GetColumnReference(cell.CellReference!);
                    columnLanguageMap[columnReference] = languageCode;
                    translations[languageCode] = new Dictionary<string, string>();

                    AppendToLog($"Found language: {languageCode} in column {columnReference}");
                }
            }

            // Process data rows
            UpdateStatusMessage("Extracting translations...");
            int resourceKeyCount = 0;
            for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var cells = row.Elements<Cell>().ToList();

                if (cells.Count == 0) continue;

                // Get resource key from first column
                string resourceKey = string.Empty;
                var keyCell = cells.FirstOrDefault();

                if (keyCell != null)
                {
                    resourceKey = GetCellValue(keyCell, sharedStringTable).Trim();
                }

                if (string.IsNullOrEmpty(resourceKey)) continue;

                resourceKeyCount++;

                // Process translation values
                foreach (var cell in cells.Skip(1))
                {
                    if (cell.CellReference == null) continue;

                    string columnRef = GetColumnReference(cell.CellReference ?? string.Empty);

                    if (columnLanguageMap.TryGetValue(columnRef, out string? languageCode))
                    {
                        string translationValue = GetCellValue(cell, sharedStringTable).Trim();

                        if (!string.IsNullOrEmpty(translationValue))
                        {
                            translations[languageCode][resourceKey] = translationValue;
                        }
                    }
                }
            }

            AppendToLog($"Processed {resourceKeyCount} resource keys for {translations.Count} languages");
        }

        // Check if there's a column explicitly marked as "default"
        bool hasDefaultColumn = translations.Keys.Any(k => k.ToLowerInvariant() == "default");
        AppendToLog($"Default language column explicitly defined: {hasDefaultColumn}");

        // Process each language
        UpdateStatusMessage("Creating/updating RESX files...");
        foreach (var languageCode in translations.Keys)
        {
            // Determine the RESX file name based on the custom prefix and language code
            string resxFileName;

            // If language code is "default" OR (it's "en" AND there's no explicit "default" column)
            if (languageCode.ToLowerInvariant() == "default" ||
                (languageCode == "en" && !hasDefaultColumn))
            {
                resxFileName = $"{resxFileNamePrefix}.resx";
                AppendToLog($"Using '{languageCode}' as the default language");
            }
            else
            {
                resxFileName = $"{resxFileNamePrefix}.{languageCode}.resx";
            }

            string resxFilePath = Path.Combine(resxFolderPath, resxFileName);

            bool fileExists = File.Exists(resxFilePath);

            AppendToLog($"Processing {resxFileName} ({translations[languageCode].Count} entries)");
            UpdateStatusMessage($"Processing {resxFileName}...");

            // Create backup if necessary
            string backupPath = string.Empty;
            if (createBackup && fileExists)
            {
                // Create a backups folder
                string backupsFolderPath = Path.Combine(resxFolderPath, "Backups");
                Directory.CreateDirectory(backupsFolderPath);

                // Create timestamped backup file
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFileName = $"{Path.GetFileNameWithoutExtension(resxFileName)}_{timestamp}{Path.GetExtension(resxFileName)}.bak";
                backupPath = Path.Combine(backupsFolderPath, backupFileName);

                File.Copy(resxFilePath, backupPath, true);
                AppendToLog($"Created backup: {backupPath}");
            }

            // Create or update the RESX file
            UpdateResxFile(resxFilePath, translations[languageCode], languageCode);

            // Add to undo stack
            if (fileExists)
            {
                if (!string.IsNullOrEmpty(backupPath))
                {
                    undoStack.Push(new UndoAction("Modify", resxFilePath, backupPath));
                    AppendToLog($"Updated: {resxFileName}");
                }
                else
                {
                    AppendToLog($"Updated: {resxFileName} (no backup created - undo unavailable)");
                }
            }
            else
            {
                undoStack.Push(new UndoAction("Create", resxFilePath));
                AppendToLog($"Created: {resxFileName}");
            }
        }

        AppendToLog("Translation processing completed successfully!");
        UpdateStatusMessage("Processing completed successfully");
    }

    private static string GetColumnReference(string? cellReference)
    {
        // Extract column reference from cell reference (e.g., "B2" => "B")
        return new string((cellReference ?? string.Empty).TakeWhile(c => !char.IsDigit(c)).ToArray());
    }

    private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
    {
        // If the cell doesn't have a value, return an empty string
        if (cell.CellValue == null)
        {
            return string.Empty;
        }

        string value = cell.CellValue.Text;

        // If the cell contains a shared string
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            // Look up the shared string
            return sharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
        }

        return value;
    }

    private static void UpdateResxFile(string resxFilePath, Dictionary<string, string> translations, string languageCode)
    {
        // If the file doesn't exist, create it with standard headers
        if (!File.Exists(resxFilePath))
        {
            using var writer = new ResXResourceWriter(resxFilePath);
            writer.AddMetadata("", "Microsoft ResX Schema");
            writer.AddMetadata("version", "2.0");
            writer.AddMetadata("reader", "System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
            writer.AddMetadata("writer", "System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");

            // Add all translations to the new file
            foreach (var kvp in translations)
            {
                writer.AddResource(kvp.Key, kvp.Value);
            }

            return;
        }

        // For existing files, we need to preserve all XML structure including comments
        // Load the existing file as XML
        XDocument doc = XDocument.Load(resxFilePath);
        XNamespace ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        // Get all existing resources and their values
        var existingResources = new Dictionary<string, XElement>();
        foreach (var dataElement in doc.Root?.Elements(ns + "data") ?? Enumerable.Empty<XElement>())
        {
            string? name = dataElement.Attribute("name")?.Value;
            if (!string.IsNullOrEmpty(name))
            {
                existingResources[name] = dataElement;
            }
        }

        // Prepare a list to track processed keys
        var processedKeys = new HashSet<string>();

        // Update existing values with new translations
        foreach (var key in existingResources.Keys.ToList())
        {
            if (translations.TryGetValue(key, out string? newValue))
            {
                // If key exists in our new translations, update the value element
                var valueElement = existingResources[key].Element(ns + "value");
                if (valueElement != null)
                {
                    valueElement.Value = newValue;
                }

                // Mark as processed
                processedKeys.Add(key);
                translations.Remove(key);
            }
        }

        // Add new translations that weren't in the original file
        foreach (var kvp in translations)
        {
            XElement newData = new XElement(ns + "data",
                new XAttribute("name", kvp.Key),
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XElement(ns + "value", kvp.Value)
            );

            doc.Root?.Add(newData);
        }

        // Save the modified XML back to the file with proper formatting
        var xws = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            NewLineChars = Environment.NewLine,
            NewLineHandling = NewLineHandling.Replace
        };

        using (var writer = XmlWriter.Create(resxFilePath, xws))
        {
            doc.Save(writer);
        }
    }

    private void ScanForExistingResxFiles()
    {
        try
        {
            // Check if folder exists
            if (!Directory.Exists(resxFolderPath))
                return;

            // Get all .resx files in the directory
            string[] resxFiles = Directory.GetFiles(resxFolderPath, "*.resx");

            if (resxFiles.Length == 0)
                return;

            // Extract the first file's name
            string fileName = Path.GetFileName(resxFiles[0]);

            // Extract the base name (either "Something.resx" or "Something.en.resx" or similar)
            string baseName;

            // Check if it's a culture-specific resx file
            if (fileName.Contains('.'))
            {
                string[] parts = fileName.Split('.');

                // If it has more than 2 parts (e.g., "Resource.en.resx")
                if (parts.Length > 2)
                {
                    baseName = parts[0];
                }
                else
                {
                    // It's a default resource file (e.g., "Resource.resx")
                    baseName = Path.GetFileNameWithoutExtension(fileName);
                }
            }
            else
            {
                // Just in case, though normally RESX files have extensions
                baseName = fileName;
            }

            // Set the prefix in the textbox
            if (!string.IsNullOrEmpty(baseName))
            {
                resxFileNamePrefix = baseName;
                txtResxFilePrefix.Text = resxFileNamePrefix;
                AppendToLog($"Detected RESX name prefix: {resxFileNamePrefix}");
            }
        }
        catch (Exception ex)
        {
            AppendToLog($"Error scanning for RESX files: {ex.Message}");
        }
    }
}