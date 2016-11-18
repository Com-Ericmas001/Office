using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using Com.Ericmas001.Common;
using Com.Ericmas001.Office.Excel.ExportParms;
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
            ExportDataTable(table, new ExcelExportParm {UseView = useView});
        }

        public void ExportDataTable(DataTable table, ExcelExportParm parms)
        {
            int totalCount = parms.UseView ? table.DefaultView.Count : table.Rows.Count;
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
                    var cols = table.Columns.OfType<DataColumn>().Select(dc => dc.ColumnName).ToList();

                    var excelHeaders = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, cols.Count]];
                    excelHeaders.Value2 = cols.ToArray();
                    excelHeaders.Font.Bold = parms.HeaderFormat.Bold;
                    excelHeaders.Interior.Color = ColorTranslator.ToOle(parms.HeaderFormat.BackColor);
                    excelHeaders.Font.Color = ColorTranslator.ToOle(parms.HeaderFormat.ForeColor);
                    var i = 2;
                    foreach (DataRow dr in parms.UseView ? table.DefaultView.OfType<DataRowView>().Select(x => x.Row) : table.Rows.OfType<DataRow>())
                    {
                        ProgressUpdated(this, new EventArgs<int>(i - 1));
                        var data = new string[cols.Count];
                        for (var j = 0; j < data.Length; ++j)
                            data[j] = dr[j].ToString();
                        var excelRow = sheet.Range[sheet.Cells[i, 1], sheet.Cells[i, cols.Count]];
                        excelRow.Value2 = data;
                        i++;
                    }
                    foreach (var colName in parms.ColumnParms.Keys)
                    {
                        var col = parms.ColumnParms[colName];
                        var colId = cols.IndexOf(colName)+1;
                        var xlCol = sheet.Range[sheet.Cells[2, colId], sheet.Cells[totalCount + 1, colId]];

                        var style = excelWorkbook.Styles.Add("col#" + colId);
                        style.HorizontalAlignment = col.HorizontalAlignment;
                        style.VerticalAlignment = col.VerticalAlignment;
                        xlCol.Style = style;

                        foreach (var fcond in col.FormatConditions)
                        {
                            ExcelApp.FormatCondition cond = xlCol.FormatConditions.Add(fcond.Type, fcond.Operator, fcond.Formula);

                            cond.Font.Bold = fcond.Format.Bold;
                            cond.Font.Color = ColorTranslator.ToOle(fcond.Format.ForeColor);
                            cond.Interior.Color = ColorTranslator.ToOle(fcond.Format.BackColor);
                        }
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
