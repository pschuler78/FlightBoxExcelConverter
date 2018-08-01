namespace FlightBoxExcelConverter
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelImportFile = new System.Windows.Forms.Label();
            this.textBoxImportFileName = new System.Windows.Forms.TextBox();
            this.buttonBrowseImportFile = new System.Windows.Forms.Button();
            this.labelExportPath = new System.Windows.Forms.Label();
            this.textBoxExportFolderName = new System.Windows.Forms.TextBox();
            this.buttonBrowseExportFolder = new System.Windows.Forms.Button();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // labelImportFile
            // 
            this.labelImportFile.AutoSize = true;
            this.labelImportFile.Location = new System.Drawing.Point(46, 43);
            this.labelImportFile.Name = "labelImportFile";
            this.labelImportFile.Size = new System.Drawing.Size(67, 13);
            this.labelImportFile.TabIndex = 0;
            this.labelImportFile.Text = "Import-Datei:";
            // 
            // textBoxImportFileName
            // 
            this.textBoxImportFileName.Location = new System.Drawing.Point(133, 40);
            this.textBoxImportFileName.Name = "textBoxImportFileName";
            this.textBoxImportFileName.Size = new System.Drawing.Size(421, 20);
            this.textBoxImportFileName.TabIndex = 1;
            // 
            // buttonBrowseImportFile
            // 
            this.buttonBrowseImportFile.Location = new System.Drawing.Point(569, 38);
            this.buttonBrowseImportFile.Name = "buttonBrowseImportFile";
            this.buttonBrowseImportFile.Size = new System.Drawing.Size(34, 23);
            this.buttonBrowseImportFile.TabIndex = 2;
            this.buttonBrowseImportFile.Text = "...";
            this.buttonBrowseImportFile.UseVisualStyleBackColor = true;
            this.buttonBrowseImportFile.Click += new System.EventHandler(this.buttonBrowseImportFile_Click);
            // 
            // labelExportPath
            // 
            this.labelExportPath.AutoSize = true;
            this.labelExportPath.Location = new System.Drawing.Point(46, 78);
            this.labelExportPath.Name = "labelExportPath";
            this.labelExportPath.Size = new System.Drawing.Size(61, 13);
            this.labelExportPath.TabIndex = 0;
            this.labelExportPath.Text = "Exportpfad:";
            // 
            // textBoxExportFolderName
            // 
            this.textBoxExportFolderName.Location = new System.Drawing.Point(133, 75);
            this.textBoxExportFolderName.Name = "textBoxExportFolderName";
            this.textBoxExportFolderName.Size = new System.Drawing.Size(421, 20);
            this.textBoxExportFolderName.TabIndex = 1;
            // 
            // buttonBrowseExportFolder
            // 
            this.buttonBrowseExportFolder.Location = new System.Drawing.Point(569, 73);
            this.buttonBrowseExportFolder.Name = "buttonBrowseExportFolder";
            this.buttonBrowseExportFolder.Size = new System.Drawing.Size(34, 23);
            this.buttonBrowseExportFolder.TabIndex = 2;
            this.buttonBrowseExportFolder.Text = "...";
            this.buttonBrowseExportFolder.UseVisualStyleBackColor = true;
            this.buttonBrowseExportFolder.Click += new System.EventHandler(this.buttonBrowseExportFolder_Click);
            // 
            // buttonConvert
            // 
            this.buttonConvert.Location = new System.Drawing.Point(49, 116);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(114, 23);
            this.buttonConvert.TabIndex = 3;
            this.buttonConvert.Text = "Konvertieren";
            this.buttonConvert.UseVisualStyleBackColor = true;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLog.Location = new System.Drawing.Point(49, 156);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLog.Size = new System.Drawing.Size(716, 282);
            this.textBoxLog.TabIndex = 4;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.buttonBrowseExportFolder);
            this.Controls.Add(this.buttonBrowseImportFile);
            this.Controls.Add(this.textBoxExportFolderName);
            this.Controls.Add(this.labelExportPath);
            this.Controls.Add(this.textBoxImportFileName);
            this.Controls.Add(this.labelImportFile);
            this.Name = "MainForm";
            this.Text = "FlightBox Excel Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelImportFile;
        private System.Windows.Forms.TextBox textBoxImportFileName;
        private System.Windows.Forms.Button buttonBrowseImportFile;
        private System.Windows.Forms.Label labelExportPath;
        private System.Windows.Forms.TextBox textBoxExportFolderName;
        private System.Windows.Forms.Button buttonBrowseExportFolder;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.TextBox textBoxLog;
    }
}

