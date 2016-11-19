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
                    ApplyStyle(excelHeaders, parms.DefaultCellFormat);
                    if(parms.HeaderFormat != null)
                        ApplyStyle(excelHeaders, parms.HeaderFormat);
                    var i = 2;
                    foreach (DataRow dr in parms.UseView ? table.DefaultView.OfType<DataRowView>().Select(x => x.Row) : table.Rows.OfType<DataRow>())
                    {
                        ProgressUpdated(this, new EventArgs<int>(i - 1));
                        var data = new string[cols.Count];
                        for (var j = 0; j < data.Length; ++j)
                            data[j] = dr[j].ToString();
                        var excelRow = sheet.Range[sheet.Cells[i, 1], sheet.Cells[i, cols.Count]];
                        ApplyStyle(excelRow, parms.DefaultCellFormat);
                        excelRow.Value2 = data;
                        excelRow.Rows.AutoFit();
                        i++;
                    }
                    foreach (var colName in parms.ColumnParms.Keys)
                    {
                        var col = parms.ColumnParms[colName];
                        var colId = cols.IndexOf(colName)+1;
                        var xlCol = sheet.Range[sheet.Cells[2, colId], sheet.Cells[totalCount + 1, colId]];

                        if (col.EnumValues != null)
                        {
                            xlCol.Validation.Add(ExcelApp.XlDVType.xlValidateList
                                , ExcelApp.XlDVAlertStyle.xlValidAlertInformation
                                , ExcelApp.XlFormatConditionOperator.xlBetween
                                , string.Join(";", col.EnumValues)
                                , Type.Missing);
                            xlCol.Validation.InCellDropdown = true;
                            xlCol.Validation.ShowError = false;
                        }
                        var style = excelWorkbook.Styles.Add("col#" + colId);
                        style.HorizontalAlignment = col.HorizontalAlignment;
                        style.VerticalAlignment = col.VerticalAlignment;
                        xlCol.Style = style;

                        foreach (var fcond in col.FormatConditions)
                        {
                            ExcelApp.FormatCondition cond = xlCol.FormatConditions.Add(fcond.Type, fcond.Operator, fcond.Formula);

                            ApplyBorderStyle(parms.DefaultCellFormat.Border, cond.Borders, true);
                            ApplyFontStyle(parms.DefaultCellFormat, cond.Font);
                            ApplyInteriorStyle(parms.DefaultCellFormat, cond.Interior);

                            ApplyBorderStyle(fcond.Format.Border, cond.Borders, true);
                            ApplyFontStyle(fcond.Format, cond.Font);
                            ApplyInteriorStyle(fcond.Format, cond.Interior);
                        }
                        ExcelApp.FormatCondition dummyCond = xlCol.FormatConditions.Add(ExcelApp.XlFormatConditionType.xlCellValue, ExcelApp.XlFormatConditionOperator.xlNotEqual, "I AM DUMB FOR DEFINING DEFAULT FORMAT");

                        ApplyBorderStyle(parms.DefaultCellFormat.Border, dummyCond.Borders, true);
                        ApplyFontStyle(parms.DefaultCellFormat, dummyCond.Font);
                        ApplyInteriorStyle(parms.DefaultCellFormat, dummyCond.Interior);

                    }
                    sheet.Application.ActiveWindow.SplitRow = 1;
                    sheet.Application.ActiveWindow.FreezePanes = true;
                    var firstRow = (ExcelApp.Range) sheet.Rows[1];
                    firstRow.AutoFilter(1, Type.Missing, ExcelApp.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    firstRow.EntireColumn.AutoFit();
                    var firstcol = (ExcelApp.Range)sheet.Columns[1];
                    firstcol.EntireRow.AutoFit();
                    excelApp.Visible = true;
                    ExportationEnded(this, new EventArgs<bool>(true));
                }
                catch
                {
                    ExportationEnded(this, new EventArgs<bool>(false));
                }
            })).Start();
        }

        private void ApplyStyle(ExcelApp.Range excelRange, ExcelFormat cellFormat)
        {
            ApplyBorderStyle(cellFormat.Border, excelRange.Borders);
            ApplyFontStyle(cellFormat, excelRange.Font);
            ApplyInteriorStyle(cellFormat, excelRange.Interior);
        }

        private void ApplyFontStyle(ExcelFormat cellFormat, ExcelApp.Font font)
        {
            font.Bold = cellFormat.Bold;

            if (cellFormat.ForeColor.HasValue)
                font.Color = ColorTranslator.ToOle(cellFormat.ForeColor.Value);
        }
        private void ApplyInteriorStyle(ExcelFormat cellFormat, ExcelApp.Interior interior)
        {
            if (cellFormat.BackColor.HasValue)
                interior.Color = ColorTranslator.ToOle(cellFormat.BackColor.Value);
        }

        private static void ApplyBorderStyle(ExcelBorder borderFormat, ExcelApp.Borders border, bool isConditionnalFormatting = false)
        {
            var sides = new[] {ExcelApp.XlBordersIndex.xlEdgeLeft, ExcelApp.XlBordersIndex.xlEdgeRight, ExcelApp.XlBordersIndex.xlEdgeTop, ExcelApp.XlBordersIndex.xlEdgeBottom, ExcelApp.XlBordersIndex.xlInsideVertical, ExcelApp.XlBordersIndex.xlInsideHorizontal};
            var condFormatSides = new[] { ExcelApp.Constants.xlLeft, ExcelApp.Constants.xlRight, ExcelApp.Constants.xlTop, ExcelApp.Constants.xlBottom }.Cast<ExcelApp.XlBordersIndex>();
            if (borderFormat != null)
            {
                foreach (var side in isConditionnalFormatting ? condFormatSides : sides)
                {
                    border[side].LineStyle = borderFormat.BorderStyle;

                    if (borderFormat.BorderColor.HasValue)
                        border[side].Color = borderFormat.BorderColor.Value;

                    if (borderFormat.BorderThickness.HasValue)
                        border[side].Weight = borderFormat.BorderThickness.Value;
                }
            }
        }
    }
}
