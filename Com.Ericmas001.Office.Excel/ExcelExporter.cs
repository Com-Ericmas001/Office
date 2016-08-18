using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using Com.Ericmas001.Common;
using DataTable = System.Data.DataTable;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel
{
    public class ExcelExporter
    {
        public event EventHandler<EventArgs<int>> ProgressUpdated = delegate { };
        public event EventHandler<EventArgs<int>> ExportationStarted = delegate { };
        public event EventHandler<EventArgs<bool>> ExportationEnded = delegate { };

        public void ExportDataTable(DataTable table, bool useView = false)
        {
            int totalCount = useView ? table.DefaultView.Count : table.Rows.Count;
            ExportationStarted(this, new EventArgs<int>(table.Rows.Count));
            new Thread(new ThreadStart(delegate
            {
                try
                {
                    var excelApp = new ExcelApp.Application();
                    var excelWorkbook = excelApp.Workbooks.Add();
                    ExcelApp.Worksheet sheet = excelWorkbook.Worksheets.Add();
                    foreach (var sh in excelWorkbook.Worksheets.Cast<ExcelApp.Worksheet>().Where(sh => sh != sheet))
                        sh.Delete();
                    sheet.Name = string.IsNullOrWhiteSpace(table.TableName) ? "Table" : table.TableName;
                    var cols = table.Columns.OfType<DataColumn>().Select(dc => dc.ColumnName).ToArray();

                    var excelHeaders = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, cols.Length]];
                    excelHeaders.Value2 = cols;
                    excelHeaders.Font.Bold = true;
                    excelHeaders.Interior.Color = ColorTranslator.ToOle(Color.Black);
                    excelHeaders.Font.Color = ColorTranslator.ToOle(Color.White);
                    var i = 2;
                    foreach (DataRow dr in useView ? table.DefaultView.OfType<DataRowView>().Select(x => x.Row) : table.Rows.OfType<DataRow>())
                    {
                        ProgressUpdated(this, new EventArgs<int>(i - 1));
                        var data = new string[cols.Length];
                        for (var j = 0; j < data.Length; ++j)
                            data[j] = dr[j].ToString();
                        var excelRow = sheet.Range[sheet.Cells[i, 1], sheet.Cells[i, cols.Length]];
                        excelRow.Value2 = data;
                        i++;
                    }
                    sheet.Application.ActiveWindow.SplitRow = 1;
                    sheet.Application.ActiveWindow.FreezePanes = true;
                    var firstRow = (ExcelApp.Range) sheet.Rows[1];
                    firstRow.AutoFilter(1, Type.Missing, ExcelApp.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    firstRow.EntireColumn.AutoFit();
                    excelApp.Visible = true;
                    ExportationEnded(this, new EventArgs<bool>(true));
                }
                catch
                {
                    ExportationEnded(this, new EventArgs<bool>(false));
                }
            })).Start();
        }
    }
}
