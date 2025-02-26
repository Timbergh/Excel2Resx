namespace Excel2Resx;

partial class MainForm
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        lblExcelFile = new Label();
        txtExcelPath = new TextBox();
        btnBrowseExcel = new Button();
        lblResxFolder = new Label();
        txtResxFolderPath = new TextBox();
        btnBrowseResxFolder = new Button();
        btnProcess = new Button();
        chkCreateBackup = new CheckBox();
        grpSettings = new GroupBox();
        lblInfoText = new Label();
        txtResxFilePrefix = new TextBox();
        lblResxFilePrefix = new Label();
        txtLog = new TextBox();
        lblLog = new Label();
        btnUndo = new Button();
        statusLabel = new Label();
        grpSettings.SuspendLayout();
        SuspendLayout();

        // Excel File Label
        lblExcelFile.AutoSize = true;
        lblExcelFile.Location = new Point(16, 30);
        lblExcelFile.Name = "lblExcelFile";
        lblExcelFile.Size = new Size(84, 20);
        lblExcelFile.TabIndex = 0;
        lblExcelFile.Text = "Excel File:";

        // Excel Path TextBox
        txtExcelPath.Location = new Point(138, 27);
        txtExcelPath.Name = "txtExcelPath";
        txtExcelPath.ReadOnly = true;
        txtExcelPath.Size = new Size(586, 27);
        txtExcelPath.TabIndex = 1;
        txtExcelPath.BackColor = Color.WhiteSmoke;
        txtExcelPath.AllowDrop = true;

        // Browse Excel Button
        btnBrowseExcel.Location = new Point(730, 26);
        btnBrowseExcel.Name = "btnBrowseExcel";
        btnBrowseExcel.Size = new Size(80, 28);
        btnBrowseExcel.TabIndex = 2;
        btnBrowseExcel.Text = "Browse...";
        btnBrowseExcel.BackColor = Color.FromArgb(240, 240, 240);
        btnBrowseExcel.UseVisualStyleBackColor = false;
        btnBrowseExcel.FlatStyle = FlatStyle.System;
        btnBrowseExcel.Click += BtnBrowseExcel_Click;

        // RESX Folder Label
        lblResxFolder.AutoSize = true;
        lblResxFolder.Location = new Point(16, 65);
        lblResxFolder.Name = "lblResxFolder";
        lblResxFolder.Size = new Size(116, 20);
        lblResxFolder.TabIndex = 3;
        lblResxFolder.Text = "RESX Folder:";

        // RESX Folder Path TextBox
        txtResxFolderPath.Location = new Point(138, 62);
        txtResxFolderPath.Name = "txtResxFolderPath";
        txtResxFolderPath.ReadOnly = true;
        txtResxFolderPath.Size = new Size(586, 27);
        txtResxFolderPath.TabIndex = 4;
        txtResxFolderPath.BackColor = Color.WhiteSmoke;

        // Browse RESX Folder Button
        btnBrowseResxFolder.Location = new Point(730, 61);
        btnBrowseResxFolder.Name = "btnBrowseResxFolder";
        btnBrowseResxFolder.Size = new Size(80, 28);
        btnBrowseResxFolder.TabIndex = 5;
        btnBrowseResxFolder.Text = "Browse...";
        btnBrowseResxFolder.BackColor = Color.FromArgb(240, 240, 240);
        btnBrowseResxFolder.UseVisualStyleBackColor = false;
        btnBrowseResxFolder.FlatStyle = FlatStyle.System;
        btnBrowseResxFolder.Click += BtnBrowseResxFolder_Click;

        // Process Button
        btnProcess.Enabled = false;
        btnProcess.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
        btnProcess.Location = new Point(710, 582);
        btnProcess.Name = "btnProcess";
        btnProcess.Size = new Size(120, 36);
        btnProcess.TabIndex = 6;
        btnProcess.Text = "Process";
        btnProcess.BackColor = Color.FromArgb(0, 120, 215);
        btnProcess.ForeColor = Color.White;
        btnProcess.UseVisualStyleBackColor = false;
        btnProcess.FlatStyle = FlatStyle.System;
        btnProcess.Click += BtnProcess_Click;

        // Create Backup Checkbox
        chkCreateBackup.AutoSize = true;
        chkCreateBackup.Checked = true;
        chkCreateBackup.CheckState = CheckState.Checked;
        chkCreateBackup.Location = new Point(400, 100);
        chkCreateBackup.Name = "chkCreateBackup";
        chkCreateBackup.Size = new Size(270, 24);
        chkCreateBackup.TabIndex = 7;
        chkCreateBackup.Text = "Create backup of existing RESX files";
        chkCreateBackup.UseVisualStyleBackColor = true;
        chkCreateBackup.CheckedChanged += ChkCreateBackup_CheckedChanged;

        // RESX Name Label
        lblResxFilePrefix.AutoSize = true;
        lblResxFilePrefix.Location = new Point(16, 100);
        lblResxFilePrefix.Name = "lblResxFilePrefix";
        lblResxFilePrefix.Size = new Size(97, 20);
        lblResxFilePrefix.TabIndex = 8;
        lblResxFilePrefix.Text = "RESX Name:";

        // RESX File Prefix TextBox
        txtResxFilePrefix.Location = new Point(138, 97);
        txtResxFilePrefix.Name = "txtResxFilePrefix";
        txtResxFilePrefix.Size = new Size(250, 27);
        txtResxFilePrefix.TabIndex = 9;
        txtResxFilePrefix.Text = "Resource";
        txtResxFilePrefix.TextChanged += TxtResxFilePrefix_TextChanged;

        // Info Text Label
        lblInfoText = new Label();
        lblInfoText.AutoSize = true;
        lblInfoText.Location = new Point(138, 127);
        lblInfoText.Name = "lblInfoText";
        lblInfoText.Size = new Size(600, 40);
        lblInfoText.Font = new Font("Segoe UI", 8.5F, FontStyle.Italic, GraphicsUnit.Point);
        lblInfoText.ForeColor = Color.FromArgb(90, 90, 90);
        lblInfoText.TabIndex = 10;
        lblInfoText.Text = "Enter the base name for RESX files (e.g. 'Resource').\r\nExisting files with this name will be modified, or new files will be created.";

        // Undo Button
        btnUndo.Enabled = false;
        btnUndo.Location = new Point(620, 582);
        btnUndo.Name = "btnUndo";
        btnUndo.Size = new Size(80, 36);
        btnUndo.TabIndex = 11;
        btnUndo.Text = "Undo";
        btnUndo.BackColor = Color.FromArgb(240, 240, 240);
        btnUndo.UseVisualStyleBackColor = false;
        btnUndo.FlatStyle = FlatStyle.System;
        btnUndo.Click += BtnUndo_Click;

        // Status Label
        statusLabel = new Label();
        statusLabel.AutoSize = true;
        statusLabel.Location = new Point(20, 588);
        statusLabel.Name = "statusLabel";
        statusLabel.Size = new Size(200, 20);
        statusLabel.TabIndex = 12;
        statusLabel.Text = "Ready";
        statusLabel.Font = new Font("Segoe UI", 9F, FontStyle.Italic, GraphicsUnit.Point);
        statusLabel.ForeColor = Color.FromArgb(90, 90, 90);

        // Settings Group Box
        grpSettings.Controls.Add(lblExcelFile);
        grpSettings.Controls.Add(txtExcelPath);
        grpSettings.Controls.Add(btnBrowseExcel);
        grpSettings.Controls.Add(lblResxFolder);
        grpSettings.Controls.Add(txtResxFolderPath);
        grpSettings.Controls.Add(btnBrowseResxFolder);
        grpSettings.Controls.Add(lblResxFilePrefix);
        grpSettings.Controls.Add(txtResxFilePrefix);
        grpSettings.Controls.Add(chkCreateBackup);
        grpSettings.Controls.Add(lblInfoText);
        grpSettings.Location = new Point(12, 12);
        grpSettings.Name = "grpSettings";
        grpSettings.Size = new Size(826, 180);
        grpSettings.TabIndex = 8;
        grpSettings.TabStop = false;
        grpSettings.Text = "Settings";

        // Log TextBox
        txtLog.BackColor = Color.WhiteSmoke;
        txtLog.Location = new Point(12, 240);
        txtLog.Multiline = true;
        txtLog.Name = "txtLog";
        txtLog.ReadOnly = true;
        txtLog.ScrollBars = ScrollBars.Vertical;
        txtLog.Size = new Size(826, 330);
        txtLog.TabIndex = 9;

        // Log Label
        lblLog.AutoSize = true;
        lblLog.Location = new Point(12, 217);
        lblLog.Name = "lblLog";
        lblLog.Size = new Size(100, 20);
        lblLog.TabIndex = 10;
        lblLog.Text = "Activity Log:";
        lblLog.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);

        // MainForm
        AutoScaleDimensions = new SizeF(8F, 20F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(850, 630);
        Controls.Add(lblLog);
        Controls.Add(txtLog);
        Controls.Add(grpSettings);
        Controls.Add(btnProcess);
        Controls.Add(btnUndo);
        Controls.Add(statusLabel);
        Name = "MainForm";
        Text = "Excel2Resx";
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterScreen;

        grpSettings.ResumeLayout(false);
        grpSettings.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Label lblExcelFile;
    private TextBox txtExcelPath;
    private Button btnBrowseExcel;
    private Label lblResxFolder;
    private TextBox txtResxFolderPath;
    private Button btnBrowseResxFolder;
    private Button btnProcess;
    private CheckBox chkCreateBackup;
    private GroupBox grpSettings;
    private TextBox txtLog;
    private Label lblLog;
    private Label lblResxFilePrefix;
    private TextBox txtResxFilePrefix;
    private Label lblInfoText;
    private Button btnUndo;
    private Label statusLabel;
}