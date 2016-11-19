using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ColumnExportParm
    {
        public XlHAlign HorizontalAlignment { get; set; }
        public XlVAlign VerticalAlignment { get; set; }
        public IEnumerable<ExcelConditionnalFormating> FormatConditions { get; set; } 
        public string[] EnumValues { get; set; }

        public ColumnExportParm()
        {
            HorizontalAlignment = XlHAlign.xlHAlignLeft;
            VerticalAlignment = XlVAlign.xlVAlignCenter;
            FormatConditions = new ExcelConditionnalFormating[0];
        }
    }
}
