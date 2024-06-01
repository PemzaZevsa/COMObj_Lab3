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
        Excel.Workbooks workbooks = null;
        Excel.Workbook workbook = null;
        Excel.Sheets sheets = null;
        Excel.Worksheet sheet = null;
        Excel.Range range = null;


        public ExcelDocument()
        {
            app = new Excel.Application();
            workbooks = app.Workbooks;
            workbook = workbooks.Add();
            sheets = app.Sheets;
            sheet = sheets[1];
        }
        public ExcelDocument(string fileName)
        {
            app = new Excel.Application();
            workbooks = app.Workbooks;
            workbook = workbooks.Add();
            sheets = app.Sheets;
            sheet = sheets[1];
        }

        public void SaveAs(string fileName)
        {
            workbook?.SaveAs(fileName);
        }

        public void Dispose()
        {
            workbook.Close();
            app?.Quit();
            Release(app);
            Release(workbooks);
            Release(workbook);
            Release(sheets);
            Release(sheet);
            Release(range);

            app = null;
        }

        public string? this[string cellName]
        {
            get
            {
                if (string.IsNullOrEmpty(cellName))
                    throw new ArgumentNullException(nameof(cellName));

                range = sheet?.Range[cellName];
                return range?.Value2?.ToString();
            }
            set
            {
                if (string.IsNullOrEmpty(cellName))
                    throw new ArgumentNullException(nameof(cellName));

                if (sheet == null)
                    throw new InvalidOperationException("Sheet is not initialized.");

                range = sheet?.Range[cellName];
                if (range != null)
                {
                    range.Value2 = value;
                }
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
