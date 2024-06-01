using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOTNET_Lab3
{
    public class ExcelDocument : IDisposable
    {
        Excel.Application app = null;
        //Excel.Workbooks workbooks = null;
        Excel.Workbook workbook = null;
        Excel.Sheets sheets = null;
        Excel.Worksheet sheet = null;
        Excel.Range range = null;


        public ExcelDocument()
        {
            app = new Excel.Application();
            //workbooks = app.Workbooks;
            //workbook = workbooks.Add();
            workbook = app.Workbooks.Add();
            sheets = app.Sheets;
            sheet = sheets[1];
        }
        public ExcelDocument(string fileName)
        {
            app = new Excel.Application();
            workbook = app.Workbooks.Open(fileName);
            sheets = app.Sheets;
            sheet = sheets[1];
        }

        public void SaveAs(string fileName)
        {
            workbook?.SaveAs(fileName);
        }

        public void AddCell(int row, int col, string text, int fontSize, bool isBold, Excel.XlRgbColor fontColor, Excel.XlHAlign Align)
        {
            range = sheet.Cells[row, col];
            range.Value = text;
            range.Font.Size = fontSize;
            range.Font.Bold = isBold;
            range.Font.Color = fontColor;
            range.HorizontalAlignment = Align;
        }

        public void AddMergedCells(int startRow, int startCol, int endRow, int endCol, string text, int fontSize, bool isBold, Excel.XlRgbColor fontColor, Excel.XlHAlign Align)
        {
            range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];

            range.Merge();

            range.Value = text;
            range.Font.Size = fontSize;
            range.Font.Bold = isBold;
            range.Font.Color = fontColor;
            range.HorizontalAlignment = Align;
        }
        public void Dispose()
        {
            workbook.Close();
            app?.Quit();
            Release(app); app = null;
            Release(workbook); workbook = null;
            Release(sheets); sheets = null;
            Release(sheet); sheet = null;
            Release(range); range = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string? this[string cellName]
        {
            get => sheet?.Range[cellName].Value2.ToString();
            set
            {
                if (sheet is not null)
                    sheet.Range[cellName].Value2 = value;
            }
        }

        public string? this[int row, int col]
        {
            get
            {
                if (sheet == null)
                    throw new InvalidOperationException("Sheet is not initialized.");

                var cell = sheet.Cells[row, col] as Excel.Range;
                return cell?.Value2?.ToString();
            }
            set
            {
                if (sheet == null)
                    throw new InvalidOperationException("Sheet is not initialized.");

                var cell = sheet.Cells[row, col] as Excel.Range;
                if (cell != null)
                {
                    cell.Value2 = value;
                }
            }
        }


        private void Release(object obj)
        {
            if (obj != null)
            {
                Marshal.FinalReleaseComObject(obj);
            }
        }
    }
}
