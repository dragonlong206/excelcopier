namespace MicrosoftExcelCopier
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.grbChoseFile = new System.Windows.Forms.GroupBox();
            this.chkSavePath = new System.Windows.Forms.CheckBox();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.ofdChooseFile = new System.Windows.Forms.OpenFileDialog();
            this.grbFunction = new System.Windows.Forms.GroupBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.chkToDate4 = new System.Windows.Forms.CheckBox();
            this.chkToDate3 = new System.Windows.Forms.CheckBox();
            this.chkToDate2 = new System.Windows.Forms.CheckBox();
            this.dtpToDate4 = new System.Windows.Forms.DateTimePicker();
            this.dtpToDate3 = new System.Windows.Forms.DateTimePicker();
            this.dtpToDate2 = new System.Windows.Forms.DateTimePicker();
            this.dtpToDate1 = new System.Windows.Forms.DateTimePicker();
            this.lblToDate = new System.Windows.Forms.Label();
            this.lblFromDate = new System.Windows.Forms.Label();
            this.dtpFromDate = new System.Windows.Forms.DateTimePicker();
            this.lblCopyright = new System.Windows.Forms.Label();
            this.grbPreview = new System.Windows.Forms.GroupBox();
            this.ecvPreviewer = new ExcelViewer.ExcelViewer();
            this.rdoPreviewOff = new System.Windows.Forms.RadioButton();
            this.rdoPreviewOn = new System.Windows.Forms.RadioButton();
            this.sfdSaveFile = new System.Windows.Forms.SaveFileDialog();
            this.grbChoseFile.SuspendLayout();
            this.grbFunction.SuspendLayout();
            this.grbPreview.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbChoseFile
            // 
            this.grbChoseFile.Controls.Add(this.chkSavePath);
            this.grbChoseFile.Controls.Add(this.txtFilePath);
            this.grbChoseFile.Controls.Add(this.btnBrowse);
            resources.ApplyResources(this.grbChoseFile, "grbChoseFile");
            this.grbChoseFile.ForeColor = System.Drawing.Color.Teal;
            this.grbChoseFile.Name = "grbChoseFile";
            this.grbChoseFile.TabStop = false;
            // 
            // chkSavePath
            // 
            resources.ApplyResources(this.chkSavePath, "chkSavePath");
            this.chkSavePath.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkSavePath.Name = "chkSavePath";
            this.chkSavePath.UseVisualStyleBackColor = true;
            // 
            // txtFilePath
            // 
            resources.ApplyResources(this.txtFilePath, "txtFilePath");
            this.txtFilePath.Name = "txtFilePath";
            // 
            // btnBrowse
            // 
            resources.ApplyResources(this.btnBrowse, "btnBrowse");
            this.btnBrowse.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnBrowse.Image = global::MicrosoftExcelCopier.Properties.Resources.browse_icon;
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // ofdChooseFile
            // 
            this.ofdChooseFile.DefaultExt = "xls;xlsx";
            resources.ApplyResources(this.ofdChooseFile, "ofdChooseFile");
            this.ofdChooseFile.FileOk += new System.ComponentModel.CancelEventHandler(this.ofdChooseFile_FileOk);
            // 
            // grbFunction
            // 
            this.grbFunction.Controls.Add(this.btnSave);
            this.grbFunction.Controls.Add(this.btnCopy);
            this.grbFunction.Controls.Add(this.chkToDate4);
            this.grbFunction.Controls.Add(this.chkToDate3);
            this.grbFunction.Controls.Add(this.chkToDate2);
            this.grbFunction.Controls.Add(this.dtpToDate4);
            this.grbFunction.Controls.Add(this.dtpToDate3);
            this.grbFunction.Controls.Add(this.dtpToDate2);
            this.grbFunction.Controls.Add(this.dtpToDate1);
            this.grbFunction.Controls.Add(this.lblToDate);
            this.grbFunction.Controls.Add(this.lblFromDate);
            this.grbFunction.Controls.Add(this.dtpFromDate);
            resources.ApplyResources(this.grbFunction, "grbFunction");
            this.grbFunction.ForeColor = System.Drawing.Color.Teal;
            this.grbFunction.Name = "grbFunction";
            this.grbFunction.TabStop = false;
            // 
            // btnSave
            // 
            resources.ApplyResources(this.btnSave, "btnSave");
            this.btnSave.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSave.Image = global::MicrosoftExcelCopier.Properties.Resources.save;
            this.btnSave.Name = "btnSave";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCopy
            // 
            resources.ApplyResources(this.btnCopy, "btnCopy");
            this.btnCopy.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnCopy.Image = global::MicrosoftExcelCopier.Properties.Resources.copy;
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // chkToDate4
            // 
            resources.ApplyResources(this.chkToDate4, "chkToDate4");
            this.chkToDate4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkToDate4.Name = "chkToDate4";
            this.chkToDate4.UseVisualStyleBackColor = true;
            this.chkToDate4.CheckedChanged += new System.EventHandler(this.chkToDate4_CheckedChanged);
            // 
            // chkToDate3
            // 
            resources.ApplyResources(this.chkToDate3, "chkToDate3");
            this.chkToDate3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkToDate3.Name = "chkToDate3";
            this.chkToDate3.UseVisualStyleBackColor = true;
            this.chkToDate3.CheckedChanged += new System.EventHandler(this.chkToDate3_CheckedChanged);
            // 
            // chkToDate2
            // 
            resources.ApplyResources(this.chkToDate2, "chkToDate2");
            this.chkToDate2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkToDate2.Name = "chkToDate2";
            this.chkToDate2.UseVisualStyleBackColor = true;
            this.chkToDate2.CheckedChanged += new System.EventHandler(this.chkToDate2_CheckedChanged);
            // 
            // dtpToDate4
            // 
            resources.ApplyResources(this.dtpToDate4, "dtpToDate4");
            this.dtpToDate4.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpToDate4.Name = "dtpToDate4";
            // 
            // dtpToDate3
            // 
            resources.ApplyResources(this.dtpToDate3, "dtpToDate3");
            this.dtpToDate3.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpToDate3.Name = "dtpToDate3";
            // 
            // dtpToDate2
            // 
            resources.ApplyResources(this.dtpToDate2, "dtpToDate2");
            this.dtpToDate2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpToDate2.Name = "dtpToDate2";
            // 
            // dtpToDate1
            // 
            resources.ApplyResources(this.dtpToDate1, "dtpToDate1");
            this.dtpToDate1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpToDate1.Name = "dtpToDate1";
            // 
            // lblToDate
            // 
            resources.ApplyResources(this.lblToDate, "lblToDate");
            this.lblToDate.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblToDate.Name = "lblToDate";
            // 
            // lblFromDate
            // 
            resources.ApplyResources(this.lblFromDate, "lblFromDate");
            this.lblFromDate.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblFromDate.Name = "lblFromDate";
            // 
            // dtpFromDate
            // 
            resources.ApplyResources(this.dtpFromDate, "dtpFromDate");
            this.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpFromDate.Name = "dtpFromDate";
            // 
            // lblCopyright
            // 
            resources.ApplyResources(this.lblCopyright, "lblCopyright");
            this.lblCopyright.Name = "lblCopyright";
            // 
            // grbPreview
            // 
            resources.ApplyResources(this.grbPreview, "grbPreview");
            this.grbPreview.Controls.Add(this.ecvPreviewer);
            this.grbPreview.Controls.Add(this.rdoPreviewOff);
            this.grbPreview.Controls.Add(this.rdoPreviewOn);
            this.grbPreview.ForeColor = System.Drawing.Color.Teal;
            this.grbPreview.Name = "grbPreview";
            this.grbPreview.TabStop = false;
            // 
            // ecvPreviewer
            // 
            resources.ApplyResources(this.ecvPreviewer, "ecvPreviewer");
            this.ecvPreviewer.ForeColor = System.Drawing.SystemColors.ControlText;
            this.ecvPreviewer.Name = "ecvPreviewer";
            // 
            // rdoPreviewOff
            // 
            resources.ApplyResources(this.rdoPreviewOff, "rdoPreviewOff");
            this.rdoPreviewOff.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rdoPreviewOff.Name = "rdoPreviewOff";
            this.rdoPreviewOff.UseVisualStyleBackColor = true;
            this.rdoPreviewOff.CheckedChanged += new System.EventHandler(this.rdoPreviewOff_CheckedChanged);
            // 
            // rdoPreviewOn
            // 
            resources.ApplyResources(this.rdoPreviewOn, "rdoPreviewOn");
            this.rdoPreviewOn.Checked = true;
            this.rdoPreviewOn.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rdoPreviewOn.Name = "rdoPreviewOn";
            this.rdoPreviewOn.TabStop = true;
            this.rdoPreviewOn.UseVisualStyleBackColor = true;
            this.rdoPreviewOn.CheckedChanged += new System.EventHandler(this.rdoPreviewOn_CheckedChanged);
            // 
            // sfdSaveFile
            // 
            this.sfdSaveFile.DefaultExt = "xls;xlsx";
            resources.ApplyResources(this.sfdSaveFile, "sfdSaveFile");
            this.sfdSaveFile.FileOk += new System.ComponentModel.CancelEventHandler(this.sfdSaveFile_FileOk);
            // 
            // frmMain
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grbPreview);
            this.Controls.Add(this.lblCopyright);
            this.Controls.Add(this.grbFunction);
            this.Controls.Add(this.grbChoseFile);
            this.Name = "frmMain";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.grbChoseFile.ResumeLayout(false);
            this.grbChoseFile.PerformLayout();
            this.grbFunction.ResumeLayout(false);
            this.grbFunction.PerformLayout();
            this.grbPreview.ResumeLayout(false);
            this.grbPreview.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grbChoseFile;
        private System.Windows.Forms.OpenFileDialog ofdChooseFile;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.GroupBox grbFunction;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.CheckBox chkSavePath;
        private System.Windows.Forms.DateTimePicker dtpFromDate;
        private System.Windows.Forms.Label lblFromDate;
        private System.Windows.Forms.Label lblToDate;
        private System.Windows.Forms.DateTimePicker dtpToDate4;
        private System.Windows.Forms.DateTimePicker dtpToDate3;
        private System.Windows.Forms.DateTimePicker dtpToDate2;
        private System.Windows.Forms.DateTimePicker dtpToDate1;
        private System.Windows.Forms.CheckBox chkToDate2;
        private System.Windows.Forms.CheckBox chkToDate4;
        private System.Windows.Forms.CheckBox chkToDate3;
        private System.Windows.Forms.Label lblCopyright;
        private System.Windows.Forms.GroupBox grbPreview;
        private System.Windows.Forms.RadioButton rdoPreviewOn;
        private System.Windows.Forms.RadioButton rdoPreviewOff;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.SaveFileDialog sfdSaveFile;
        private ExcelViewer.ExcelViewer ecvPreviewer;
    }
}

