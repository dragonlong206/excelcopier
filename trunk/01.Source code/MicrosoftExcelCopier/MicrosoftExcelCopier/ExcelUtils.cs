using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using System.Globalization;
using System.IO;

namespace MicrosoftExcelCopier
{
    public static class ExcelUtils
    {
        /// <summary>
        /// Get excel cell's value
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="formulaEvaluator"></param>
        /// <returns></returns>
        public static object GetCellValue(ICell cell, IFormulaEvaluator formulaEvaluator)
        {
            return GetCellValue(cell, formulaEvaluator, true);
        }

        /// <summary>
        /// Get excel cell's value
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="formulaEvaluator"></param>
        /// <param name="ignoreSumFormula">If true, this method returns cell value which is evaluated from formula when cell type is formula. Else, this method return formula of SUM formula and value of other formulas.</param>
        /// <returns></returns>
        public static object GetCellValue(ICell cell, IFormulaEvaluator formulaEvaluator, bool ignoreSumFormula)
        {
            if (null == cell)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.BOOLEAN:
                    return cell.BooleanCellValue;
                case CellType.ERROR:
                    return cell.ErrorCellValue;
                case CellType.FORMULA:
                    // If not ignore sum formula then return the formula
                    if (!ignoreSumFormula && cell.CellFormula.Contains(Properties.Settings.Default.SUMFormula))
                    {
                        return cell.CellFormula;                        
                    }
                    else
                    {
                        CellValue cellValue = formulaEvaluator.Evaluate(cell);
                        switch (cellValue.CellType)
                        {
                            case CellType.BOOLEAN:
                                return cellValue.BooleanValue;
                            case CellType.ERROR:
                                return cellValue.ErrorValue;
                            case CellType.NUMERIC:
                                return cellValue.NumberValue;
                            case CellType.STRING:
                                return cellValue.StringValue;
                            default:
                                return cellValue.StringValue;
                        }
                    }
                    //return dataFormatter.FormatCellValue(cell, formulaEvaluator);
                case CellType.NUMERIC:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString(Properties.Settings.Default.DateFormat);
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.STRING:
                    return cell.StringCellValue.Trim();
                default:
                    return cell.StringCellValue;
            }
        }

        /// <summary>
        /// Get cell's comment
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetCellComment(ICell cell)
        {
            string ret = string.Empty;
            if ((null == cell) || (null == cell.CellComment)) { return ret; }
            IRichTextString str = cell.CellComment.String;
            if (str != null && str.Length > 0)
            {
                ret = str.ToString();
            }
            return ret;
        }

        /// <summary>
        /// Return merged region has value matching to "value". If it doesn't exist, return null.
        /// </summary>
        /// <param name="value">Value to find</param>
        /// <returns></returns>
        public static CellRangeAddress FindMergedRegion(this ISheet sheet, string value)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException("sheet");
            }

            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            int numMergedRegion = sheet.NumMergedRegions;
            IFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(sheet.Workbook);
            for (int i = 0; i < numMergedRegion; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);

                if (region != null)
                {
                    // Traversal each cell in region
                    for (int j = region.FirstRow; j <= region.LastRow; j++)
                    {
                        IRow row = sheet.GetRow(j);
                        int numberOfColumn = region.LastColumn - region.FirstColumn + 1;
                        for (int k = 0; k < numberOfColumn; k++)
                        {
                            ICell cell = row.Cells[k];
                            object cellValue = GetCellValue(cell, formulaEvaluator);
                            if (value.Equals(cellValue))
                            {
                                return region;
                            }
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get data in stock row. If stock row does not exist, return empty list.
        /// </summary>
        /// <param name="sheet">Data sheet</param>
        /// <param name="region">Source date region</param>
        /// <param name="stockType">Type of stock: STOCK, STOCKVN or STOCKTQ</param>
        /// <returns></returns>
        public static List<object> GetData(this ISheet sheet, CellRangeAddress region, StockType stockType)
        {
            List<object> result = new List<object>();

            string label = string.Empty;
            switch (stockType)
            {
                case (StockType.STOCK):
                    label = Properties.Settings.Default.StockLabel;
                    break;
                case (StockType.STOCKTQ):
                    label = Properties.Settings.Default.StockTQLabel;
                    break;
                case (StockType.STOCKVN):
                    label = Properties.Settings.Default.StockVNLabel;
                    break;
            }

            IFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(sheet.Workbook);
            for (int i = region.LastRow; i >= region.FirstRow; i--)
            {
                IRow row = sheet.GetRow(i);
                ICell labelCell = row.GetCell(Properties.Settings.Default.LabelColumnNumber);
                object cellValue = GetCellValue(labelCell, formulaEvaluator);
                
                if (cellValue != null && label.Equals(cellValue.ToString(), StringComparison.InvariantCultureIgnoreCase))
                {
                    for (int j = Properties.Settings.Default.LabelColumnNumber + 1; j < row.LastCellNum; j++)
                    {
                        result.Add(GetCellValue(row.GetCell(j), formulaEvaluator, false));
                    }

                    // Total cell
                    //result.Add(GetCellValue(row.GetCell(row.LastCellNum), formulaEvaluator, false));

                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Write data to "Opening" row. If stock type is "Stock" then write to "Opening" row. If stock type is "StockVN" then write to "OpeningVN" row. If stock type is "StockTQ" then write to "OpeningTQ" row.
        /// </summary>
        /// <param name="sheet">Sheet contains data</param>
        /// <param name="data">Data to write, copied from stock row.</param>
        /// <param name="region">Date region</param>
        /// <param name="stockType">Stock type: Stock, StockVN, StockTQ</param>
        public static void WriteOpeningData(this ISheet sheet, List<object> data, CellRangeAddress region, StockType stockType)
        {
            if (data != null && data.Count > 0)
            {
                string label = string.Empty;
                switch (stockType)
                {
                    case (StockType.STOCK):
                        label = Properties.Settings.Default.OpeningLabel;
                        break;
                    case (StockType.STOCKTQ):
                        label = Properties.Settings.Default.OpeningTQLabel;
                        break;
                    case (StockType.STOCKVN):
                        label = Properties.Settings.Default.OpeningVNLabel;
                        break;
                }

                ICell cell;
                double value = 0;
                string stringValue = string.Empty;
                //DateTime datetimeValue;
                int labelDistance = Properties.Settings.Default.LabelColumnNumber + 1; // Number of label columns
                int lastCellNum = 0;
                IFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(sheet.Workbook);
                for (int i = region.LastRow; i >= region.FirstRow; i--)
                {
                    IRow row = sheet.GetRow(i);
                    ICell labelCell = row.GetCell(Properties.Settings.Default.LabelColumnNumber);
                    object cellValue = GetCellValue(labelCell, formulaEvaluator);
                    // "Opening" row
                    if (cellValue != null && label.Equals(cellValue.ToString(), StringComparison.InvariantCultureIgnoreCase))
                    {
                        lastCellNum = Math.Min(row.LastCellNum, labelDistance + data.Count);

                        for (int j = labelDistance; j < lastCellNum; j++)
                        {
                            cell = row.GetCell(j);
                            if (cell == null)
                            {
                                cell = row.CreateCell(j);
                            }

                            if (data[j - labelDistance] != null)
                            {
                                stringValue = data[j - labelDistance].ToString();
                            }
                            else
                            {
                                stringValue = string.Empty;
                            }

                            if (!string.IsNullOrEmpty(stringValue) && !stringValue.Contains(Properties.Settings.Default.SUMFormula)) // Ignore SUM cell
                            {
                                if (double.TryParse(stringValue, out value))
                                {
                                    cell.SetCellType(CellType.NUMERIC);
                                    cell.SetCellValue(value);
                                }
                                else // TODO: datetime value
                                {
                                    // Formula
                                    if (stringValue.StartsWith(Properties.Settings.Default.StartFormulaSymbol))
                                    {
                                        cell.SetCellType(CellType.FORMULA);
                                        cell.SetCellFormula(stringValue);
                                    }
                                    else
                                    {
                                        cell.SetCellType(CellType.STRING);
                                        cell.SetCellValue(stringValue);
                                    }
                                }
                            }
                        }

                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Copy all things to new excel file. Add or remove date (according to month) and clear data.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="filePath"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns>Success or fail</returns>
        public static bool CopyFile(IWorkbook workbook, string filePath, DateTime fromDate, DateTime toDate)
        {
            LogServices.WriteDebug(string.Format("Copy from workbook to new file {0}, {1} to {2}", filePath, fromDate.ToString(), toDate.ToString()));

            int monthDistance = toDate.Month - fromDate.Month;

            try
            {
                // First, copy old workbook to new file
                using (FileStream stream = File.Open(filePath, FileMode.OpenOrCreate))
                {
                    workbook.Write(stream);
                    //LogServices.WriteDebug("1. Create new file and copy all workbook to the file");
                }

                // Second, modify data

                int daysInFromMonth = DateTime.DaysInMonth(fromDate.Year, fromDate.Month);
                int daysInToMonth = DateTime.DaysInMonth(toDate.Year, toDate.Month);

                IWorkbook tempWorkbook;
                using (FileStream stream = File.Open(filePath, FileMode.Open))
                {
                    tempWorkbook = new HSSFWorkbook(stream);
                    if (tempWorkbook != null)
                    {
                        HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(tempWorkbook);
                        #region Copy and clear all data
                        DateTime oldDate;
                        bool hasRedudantRow = false; ; // Flag to know redudant row if new month has less day then old month
                        int redudantRowIndex = -1;
                        foreach (ISheet sheet in tempWorkbook)
                        {
                            redudantRowIndex = -1;
                            hasRedudantRow = false;
                            foreach (IRow row in sheet)
                            {
                                if (!hasRedudantRow || (hasRedudantRow && row.RowNum < redudantRowIndex))
                                {
                                    foreach (ICell cell in row)
                                    {
                                        // Copy value
                                        object cellValue;
                                        if (cell.CellType == CellType.NUMERIC)
                                        {
                                            if (DateUtil.IsCellDateFormatted(cell))
                                            {
                                                oldDate = cell.DateCellValue;
                                                if (oldDate != null)
                                                {
                                                    cellValue = oldDate.AddMonths(monthDistance);
                                                    cell.SetCellValue((DateTime)cellValue);
                                                    
                                                    if (oldDate.IsLastDayOfMonth()) // if last day then all remain cell is redudant)
                                                    {
                                                        LogServices.WriteDebug("Has redudant date from " + oldDate.ToString());
                                                        hasRedudantRow = true;
                                                        CellRangeAddress dateRegion = sheet.FindMergedRegion(((DateTime)cellValue).ToString(Properties.Settings.Default.DateFormat));
                                                        if (dateRegion != null)
                                                        {
                                                            redudantRowIndex = dateRegion.LastRow + 1;
                                                            LogServices.WriteDebug("Find redudant region start from row " + redudantRowIndex.ToString());
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                cell.SetCellValue(0);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //sheet.RemoveRow(row);
                                    break;
                                }
                            }

                            // Remove all redudant rows
                            if (hasRedudantRow && redudantRowIndex > 0)
                            {
                                LogServices.WriteDebug("Remove redudant rows");

                                List<int> redudantMergedRegionIndices = new List<int>();
                                for (int j = 0; j < sheet.NumMergedRegions; j++)
                                {
                                    CellRangeAddress region = sheet.GetMergedRegion(j);
                                    if (region != null && region.FirstRow >= redudantRowIndex)
                                    {
                                        redudantMergedRegionIndices.Add(j);
                                    }
                                }

                                if (redudantMergedRegionIndices.Count > 0)
                                {
                                    LogServices.WriteDebug("Remove all redudant merged regions: " + string.Join(", ", redudantMergedRegionIndices));
                                    foreach (int index in redudantMergedRegionIndices)
                                    {
                                        sheet.RemoveMergedRegion(index);
                                    }
                                }

                                LogServices.WriteDebug("Remove all rows from " + redudantRowIndex.ToString());
                                for (int i = redudantRowIndex; i <= sheet.LastRowNum; i++)
                                {
                                    IRow row = sheet.GetRow(i);
                                    if (row != null)
                                    {
                                        sheet.RemoveRow(row);
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Insert new row if to month has more day then from month
                        if (daysInFromMonth < daysInToMonth)
                        {
                            DateTime lastDateOfMonth = fromDate.GetLastDayOfMonth().AddMonths(monthDistance);
                            DateTime currentDate;
                            foreach (ISheet sheet in tempWorkbook)
                            {
                                currentDate = lastDateOfMonth.AddDays(1);
                                CellRangeAddress lastDateRegion = sheet.FindMergedRegion(lastDateOfMonth.ToString(Properties.Settings.Default.DateFormat));
                                if (lastDateRegion != null)
                                {
                                    int numberOfRegionRow = lastDateRegion.LastRow - lastDateRegion.FirstRow + 1; // 0-based
                                    int dayDistance = daysInToMonth - daysInFromMonth;
                                    for (int i = 0; i < dayDistance; i++)
                                    {
                                        int lastRow = sheet.LastRowNum + 1;
                                        for (int j = 0; j < numberOfRegionRow; j++)
                                        {
                                            IRow row = sheet.CreateRow(lastRow + j);
                                            IRow sourceRow = sheet.GetRow(lastDateRegion.FirstRow + j);
                                            int colNum = 0;
                                            foreach (ICell cell in sourceRow)
                                            {
                                                ICell newCell = row.CreateCell(colNum, cell.CellType);
                                                newCell.CellStyle = cell.CellStyle;

                                                switch (cell.CellType)
                                                {
                                                    case CellType.FORMULA:
                                                        newCell.CellFormula = cell.CellFormula;
                                                        break;
                                                    case CellType.NUMERIC:
                                                        if (DateUtil.IsCellDateFormatted(cell))
                                                        {
                                                            newCell.SetCellValue(currentDate);
                                                        }
                                                        else
                                                        {
                                                            newCell.SetCellValue(cell.NumericCellValue);
                                                        }
                                                        break;
                                                    case CellType.STRING:
                                                        newCell.SetCellValue(cell.StringCellValue);
                                                        break;
                                                    default:
                                                        break;
                                                }

                                                colNum++;
                                            }

                                            CellRangeAddress newMergedRegion = new CellRangeAddress(lastRow, lastRow + numberOfRegionRow - 1, lastDateRegion.FirstColumn, lastDateRegion.LastColumn);
                                            sheet.AddMergedRegion(newMergedRegion);
                                        }

                                        currentDate = currentDate.AddDays(1);
                                    }
                                }
                            }
                        }
                        #endregion

                        // Re-calculate all formulas
                        HSSFFormulaEvaluator.EvaluateAllFormulaCells(tempWorkbook);
                    }
                }

                using (FileStream stream = File.Open(filePath, FileMode.OpenOrCreate))
                {
                    if (tempWorkbook != null)
                    {
                        tempWorkbook.Write(stream);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                LogServices.WriteError("Khong the tao tap tin cho thang moi", ex);
                return false;
            }
        }

        public static void ClearCell(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.NUMERIC:
                    cell.SetCellValue(0);
                    break;
                case CellType.STRING:
                    cell.SetCellValue(string.Empty);
                    break;
                case CellType.BOOLEAN:
                    cell.SetCellValue(false);
                    break;
                default:
                    break;
            }
        }

        public static string GetCellAddress(int rowIndex, int columnIndex)
        {
            string result = string.Empty;

            if (columnIndex > Properties.Settings.Default.NumberOfLetter)
            {
                result += (char)(columnIndex / Properties.Settings.Default.NumberOfLetter + 65); // 65: A
            }

            result += (char)(columnIndex % Properties.Settings.Default.NumberOfLetter + 65); // 65: A

            result += rowIndex.ToString();

            return result;
        }
    }
}
