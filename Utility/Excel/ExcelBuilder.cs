using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Utility.Excel
{
    /// <summary>
    /// Utility class for building Excel files using EPPlus.
    /// Supports creating multiple sheets with configurable columns and rows.
    /// </summary>
    public static class ExcelBuilder
    {
        #region Create Methods

        /// <summary>
        /// Creates a new Excel file with a single sheet.
        /// </summary>
        /// <param name="filePath">Output file path.</param>
        /// <param name="columns">Column definitions (header + cell configuration).</param>
        /// <param name="items">Row data items.</param>
        public static void CreateNew(string filePath, List<ExcelColumn> columns, List<IExcelRow> items)
        {
            CreateNew(filePath, new SheetConfig { Columns = columns, Items = items });
        }

        /// <summary>
        /// Creates a new Excel file with one or more sheet configurations.
        /// </summary>
        /// <param name="filePath">Output file path.</param>
        /// <param name="sheets">Sheet configurations.</param>
        public static void CreateNew(string filePath, params SheetConfig[] sheets)
        {
            using (ExcelPackage xlPackage = Create(sheets))
            {
                string dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                xlPackage.SaveAs(new FileInfo(filePath));
            }
        }

        /// <summary>
        /// Creates an ExcelPackage object in memory without saving to disk.
        /// </summary>
        /// <param name="sheets">Sheet configurations.</param>
        /// <returns>ExcelPackage instance.</returns>
        public static ExcelPackage Create(params SheetConfig[] sheets)
        {
            ExcelPackage xlPackage = new ExcelPackage();

            foreach (SheetConfig sheet in sheets)
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(sheet.Name);
                
                FillHead(worksheet, sheet.Columns.Select(c => c.HeadSettings));
                FillItems(worksheet, 2, sheet.Columns.Select(c => c.CellSettings), sheet.Items);

                if (sheet.AutoFit)
                {
                    for (int i = 1; i <= sheet.Columns.Count; i++)
                    {
                        worksheet.Column(i).AutoFit();
                    }
                }
            }

            return xlPackage;
        }

        #endregion

        #region Fill Methods

        /// <summary>
        /// Fills the header row (row 1) with column names and styles.
        /// </summary>
        private static void FillHead(ExcelWorksheet sheet, IEnumerable<ExcelCell> cellSettings)
        {
            FillRow(sheet, 1, cellSettings);
        }

        /// <summary>
        /// Fills multiple rows of data starting at the specified index.
        /// </summary>
        private static void FillItems(ExcelWorksheet sheet, int startRowIndex, IEnumerable<ExcelCell> cellSettings, List<IExcelRow> items)
        {
            if (items == null || items.Count == 0) return;

            int rowIndex = startRowIndex;
            foreach (IExcelRow item in items)
            {
                FillRow(sheet, rowIndex, cellSettings, item); rowIndex++;
            }
        }

        /// <summary>
        /// Fills a single row with values or headers.
        /// </summary>
        private static void FillRow(ExcelWorksheet sheet, int rowIndex, IEnumerable<ExcelCell> cellSettings, IExcelRow item = null)
        {
            if(cellSettings == null || cellSettings.Count()== 0) return;

            int colIndex = 1;
            foreach (ExcelCell setting in cellSettings)
            {
                ExcelRange cell = sheet.Cells[rowIndex, colIndex];
                FillCell(cell, setting, item);
                colIndex++;
            }
        }

        /// <summary>
        /// Fills one cell with either header text or item value, and applies cell style if available.
        /// </summary>
        private static void FillCell(ExcelRange cell, ExcelCell settings, IExcelRow item = null)
        {
            // If no row item is provided, treat as header cell
            cell.Value = item == null ? settings.FieldName : settings.GetValue(item);
            // Apply style if defined
            settings.Style?.SetCell(cell);
        }

        #endregion
    }
}
