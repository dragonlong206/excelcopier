using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Util;

namespace MicrosoftExcelCopier
{
    public partial class frmMain : Form
    {
        private IWorkbook workbook;
        private IWorkbook destWorkbook;
        private bool isSaveResult = false;
        private bool isCopySuccess = false;
        private bool isSameMonth = true;

        private string sourceFile;

        public frmMain()
        {
            CultureInfo currentCulture = new CultureInfo(Properties.Settings.Default.Culture);
            Thread.CurrentThread.CurrentCulture = currentCulture;
            Thread.CurrentThread.CurrentUICulture = currentCulture;

            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            ofdChooseFile.ShowDialog(this);
        }

        private void ofdChooseFile_FileOk(object sender, CancelEventArgs e)
        {
            if (File.Exists(ofdChooseFile.FileName))
            {
                this.txtFilePath.Text = ofdChooseFile.FileName;
                this.chkSavePath.Checked = true;

                ShowPreview(ofdChooseFile.FileName);
                this.btnSave.Enabled = false;
                this.sourceFile = ofdChooseFile.FileName;

                // Mark new file
                this.isCopySuccess = false;
            }
            else
            {
                FormUtil.ShowMessageBoxLocalize(this, vi_VN.fileNotExist, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            //if (!string.IsNullOrEmpty(Properties.Settings.Default.DefaultFilePath))
            //{
            //    // Check file exists
            //    if (File.Exists(Properties.Settings.Default.DefaultFilePath))
            //    {
            //        this.txtFilePath.Text = Properties.Settings.Default.DefaultFilePath;

            //        //LoadExcelFile(Properties.Settings.Default.DefaultFilePath);
            //    }
            //}

            // Default date
            this.dtpFromDate.Value = DateTime.Now.Subtract(new TimeSpan(1, 0, 0, 0));
            this.dtpToDate1.Value = DateTime.Now;
            this.dtpToDate2.Value = DateTime.Now.AddDays(1);
            this.dtpToDate3.Value = DateTime.Now.AddDays(2);
            this.dtpToDate4.Value = DateTime.Now.AddDays(3);

            // Preview
            if (Properties.Settings.Default.Preview)
            {
                this.rdoPreviewOn.Checked = true;
                this.rdoPreviewOff.Checked = false;
            }
            else
            {
                this.rdoPreviewOn.Checked = false;
                this.rdoPreviewOff.Checked = true;
            }
        }

        private void chkToDate2_CheckedChanged(object sender, EventArgs e)
        {
            this.dtpToDate2.Enabled = this.chkToDate2.Checked;
        }

        private void chkToDate3_CheckedChanged(object sender, EventArgs e)
        {
            this.dtpToDate3.Enabled = this.chkToDate3.Checked;
        }

        private void chkToDate4_CheckedChanged(object sender, EventArgs e)
        {
            this.dtpToDate4.Enabled = this.chkToDate4.Checked;
        }

        /// <summary>
        /// Load excel file and display on viewer
        /// </summary>
        /// <param name="filePath">Full path of file</param>
        private void ShowPreview(string filePath)
        {
            if (rdoPreviewOn.Checked)
            {
                if (File.Exists(filePath))
                {
                    ecvPreviewer.CloseExcel();
                    ecvPreviewer.OpenFile(filePath);
                }
                //else
                //{
                //    FormUtil.ShowMessageBoxLocalize(this, vi_VN.fileNotExist, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //}
            }
        }

        private void rdoPreviewOn_CheckedChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Preview = true;
            ecvPreviewer.Show();
            if (!this.isCopySuccess)
            {
                ShowPreview(this.sourceFile);
            }
            else
            {
                if (this.isSameMonth)
                {
                    ShowPreview(Path.GetFullPath(Properties.Settings.Default.TempFile));
                }
                else
                {
                    ShowPreview(Path.GetFullPath(Properties.Settings.Default.NewMonthTempFile));
                }
            }
        }

        private void rdoPreviewOff_CheckedChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Preview = false;
            //ecvPreviewer.CloseExcel();
            ecvPreviewer.Hide();
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.btnSave.Enabled && !this.isSaveResult)
            {
                DialogResult choice = FormUtil.ShowMessageBoxLocalize(this, vi_VN.saveQuestion, vi_VN.guideCaption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (choice == System.Windows.Forms.DialogResult.Yes)
                {
                    btnSave_Click(sender, e);
                }
                else if (choice == System.Windows.Forms.DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            ecvPreviewer.CloseExcel();

            if (!string.IsNullOrEmpty(sourceFile))
            {
                DateTime fromDate = dtpFromDate.Value;
                string fromDateFormatted = fromDate.ToString(Properties.Settings.Default.DateFormat);
                LogServices.WriteDebug(string.Format("Copy data from {0}", fromDateFormatted));

                DateTime toDate1 = dtpToDate1.Value;
                string toDate1Formatted = toDate1.ToString(Properties.Settings.Default.DateFormat);

                string toDate2Formatted = string.Empty;
                if (chkToDate2.Checked)
                {
                    // Check to dates in a month
                    if (dtpToDate2.Value.Month != toDate1.Month)
                    {
                        FormUtil.ShowMessageBoxLocalize(this, vi_VN.toDateInAMonth, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    toDate2Formatted = dtpToDate2.Value.ToString(Properties.Settings.Default.DateFormat);
                }

                string toDate3Formatted = string.Empty;
                if (chkToDate3.Checked)
                {
                    // Check to dates in a month
                    if (dtpToDate3.Value.Month != toDate1.Month)
                    {
                        FormUtil.ShowMessageBoxLocalize(this, vi_VN.toDateInAMonth, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    toDate3Formatted = dtpToDate3.Value.ToString(Properties.Settings.Default.DateFormat);
                }

                string toDate4Formatted = string.Empty;
                if (chkToDate4.Checked)
                {
                    // Check to dates in a month
                    if (dtpToDate4.Value.Month != toDate1.Month)
                    {
                        FormUtil.ShowMessageBoxLocalize(this, vi_VN.toDateInAMonth, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    toDate4Formatted = dtpToDate4.Value.ToString(Properties.Settings.Default.DateFormat);
                }

                isSameMonth = true;
                if (fromDate.Month != toDate1.Month)
                {
                    isSameMonth = false;
                    //LogServices.WriteDebug("Different month");
                }
                try
                {
                    // Flag to know there are data of selected date or not
                    //bool hasData = false;
                    List<ISheet> dataNotFoundSheet = new List<ISheet>();

                    using (StreamReader input = new StreamReader(sourceFile))
                    {
                        //LogServices.WriteDebug(string.Format("Start openning file: {0}", sourceFile));

                        workbook = new HSSFWorkbook(new POIFSFileSystem(input.BaseStream));
                        destWorkbook = workbook;
                        if (null == workbook)
                        {
                            LogServices.WriteError(string.Format("Cannot open file: {0}", sourceFile));

                            FormUtil.ShowMessageBoxLocalize(this, string.Format(vi_VN.cannotOpenFile, sourceFile), vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            return;
                        }

                        if (!isSameMonth)
                        {
                            bool createNewFileSuccess = ExcelUtils.CopyFile(workbook, Properties.Settings.Default.NewMonthTempFile, fromDate, toDate1);
                            if (!createNewFileSuccess)
                            {
                                FormUtil.ShowMessageBoxLocalize(this, vi_VN.cannotCreateFile, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }

                        //IFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(workbook);
                        //DataFormatter dataFormatter = new HSSFDataFormatter(new CultureInfo("en-US"));

                        string value = string.Empty;
                        if (!isSameMonth)
                        {
                            using (FileStream destStream = File.Open(Properties.Settings.Default.NewMonthTempFile, FileMode.Open))
                            {
                                destWorkbook = new HSSFWorkbook(destStream);
                            }
                        }

                        foreach (ISheet sheet in workbook)
                        {
                            CellRangeAddress fromDateRegion = sheet.FindMergedRegion(fromDateFormatted);
                            ISheet destSheet = sheet;
                            if (!isSameMonth)
                            {
                                destSheet = destWorkbook.GetSheet(sheet.SheetName);
                            }

                            if (fromDateRegion != null)
                            {
                                StockType stockType = StockType.STOCK;
                                CopyAndPasteAllDates(sheet, destSheet, stockType, fromDateRegion, toDate1Formatted, toDate2Formatted, toDate3Formatted, toDate4Formatted);

                                stockType = StockType.STOCKTQ;
                                CopyAndPasteAllDates(sheet, destSheet, stockType, fromDateRegion, toDate1Formatted, toDate2Formatted, toDate3Formatted, toDate4Formatted);

                                stockType = StockType.STOCKVN;
                                CopyAndPasteAllDates(sheet, destSheet, stockType, fromDateRegion, toDate1Formatted, toDate2Formatted, toDate3Formatted, toDate4Formatted);
                            }
                            else
                            {
                                // Data not found
                                dataNotFoundSheet.Add(sheet);
                            }
                        }
                    }

                    // All sheet do not have data
                    if (dataNotFoundSheet.Count == workbook.NumberOfSheets)
                    {
                        FormUtil.ShowMessageBoxLocalize(this, string.Format(vi_VN.dataNotFound, fromDateFormatted), vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        if (destWorkbook != null)
                        {
                            // Re-calculate all formula
                            HSSFFormulaEvaluator.EvaluateAllFormulaCells(destWorkbook);

                            using (FileStream stream = File.Open(Properties.Settings.Default.TempFile, FileMode.OpenOrCreate))
                            {
                                destWorkbook.Write(stream);
                            }

                            ShowPreview(Path.GetFullPath(Properties.Settings.Default.TempFile));
                        }

                        if (dataNotFoundSheet.Count > 0)
                        {
                            string dataNotFoundSheetNames = string.Join(", ", dataNotFoundSheet.Select(x => x.SheetName));
                            FormUtil.ShowMessageBoxLocalize(this, string.Format(vi_VN.dataNotFoundInSomeSheet, fromDateFormatted, dataNotFoundSheetNames), vi_VN.infoCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            FormUtil.ShowMessageBoxLocalize(this, vi_VN.copySucessfully, vi_VN.infoCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        this.isSaveResult = false;
                        this.btnSave.Enabled = true;
                        this.isCopySuccess = true;
                        //LogServices.WriteDebug(string.Format("Closing file: {0}", sourceFile));
                    }
                }
                catch (IOException ioEx)
                {
                    LogServices.WriteError(string.Format("Cannot open file: {0}", sourceFile), ioEx);

                    FormUtil.ShowMessageBoxLocalize(this, string.Format(vi_VN.cannotOpenFile, sourceFile), vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                catch (Exception ex)
                {
                    LogServices.WriteError("Exception: ", ex);
                    FormUtil.ShowMessageBoxLocalize(this, vi_VN.systemError, vi_VN.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    this.Cursor = this.DefaultCursor;
                }
            }
            else
            {
                FormUtil.ShowMessageBoxLocalize(this, vi_VN.filePathEmpty, vi_VN.guideCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            this.Cursor = this.DefaultCursor;
        }

        /// <summary>
        /// Write data to a specific date
        /// </summary>
        /// <param name="toDate"></param>
        /// <param name="sheet"></param>
        /// <param name="copyData"></param>
        /// <param name="stockType"></param>
        private void PasteToOneDate(string toDate, ISheet sheet, List<object> copyData, StockType stockType)
        {
            CellRangeAddress toDateRegion = sheet.FindMergedRegion(toDate);
            if (toDateRegion != null)
            {
                LogServices.WriteDebug(string.Format("Write data: \"{0}\" to day {1}, stock type {2}", string.Join(", ", copyData), toDate, stockType));
                sheet.WriteOpeningData(copyData, toDateRegion, stockType);
            }
            else
            {
                LogServices.WriteError(string.Format("Cannot find date {0} in sheet {1}, file {2}", toDate, sheet.SheetName, sourceFile));
            }
        }

        /// <summary>
        /// Read and write data to all selected dates
        /// </summary>
        /// <param name="sourceSheet"></param>
        /// <param name="destSheet"></param>
        /// <param name="stockType"></param>
        /// <param name="fromDateRegion"></param>
        /// <param name="toDate1Formatted"></param>
        /// <param name="toDate2Formatted"></param>
        /// <param name="toDate3Formatted"></param>
        /// <param name="toDate4Formatted"></param>
        /// <returns>Have data or not</returns>
        private bool CopyAndPasteAllDates(ISheet sourceSheet, ISheet destSheet, StockType stockType, CellRangeAddress fromDateRegion, string toDate1Formatted, string toDate2Formatted, string toDate3Formatted, string toDate4Formatted)
        {
            if (sourceSheet != null && destSheet != null)
            {
                List<object> copyData = sourceSheet.GetData(fromDateRegion, stockType);

                if (copyData != null && copyData.Count > 0)
                {
                    PasteToOneDate(toDate1Formatted, destSheet, copyData, stockType);
                    if (!string.IsNullOrEmpty(toDate2Formatted))
                    {
                        PasteToOneDate(toDate2Formatted, destSheet, copyData, stockType);
                    }
                    if (!string.IsNullOrEmpty(toDate3Formatted))
                    {
                        PasteToOneDate(toDate3Formatted, destSheet, copyData, stockType);
                    }
                    if (!string.IsNullOrEmpty(toDate4Formatted))
                    {
                        PasteToOneDate(toDate4Formatted, destSheet, copyData, stockType);
                    }
                    return true;
                }
            }
            return false;
        }

        private bool CopyAndPasteAllDates(ISheet sheet, StockType stockType, CellRangeAddress fromDateRegion, string toDate1Formatted, string toDate2Formatted, string toDate3Formatted, string toDate4Formatted)
        {
            return CopyAndPasteAllDates(sheet, sheet, stockType, fromDateRegion, toDate1Formatted, toDate2Formatted, toDate3Formatted, toDate4Formatted);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.sfdSaveFile.FileName = Path.GetFileName(sourceFile);
            this.sfdSaveFile.ShowDialog(this);
        }

        private void sfdSaveFile_FileOk(object sender, CancelEventArgs e)
        {
            SaveResult();
        }

        private void SaveResult()
        {
            string destFile = this.sfdSaveFile.FileName;
            if (!string.IsNullOrEmpty(destFile))
            {
                using (FileStream stream = File.Open(destFile, FileMode.OpenOrCreate))
                {
                    destWorkbook.Write(stream);
                }

                this.isSaveResult = true;
                FormUtil.ShowMessageBoxLocalize(this, vi_VN.saveSuccessfully, vi_VN.infoCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
