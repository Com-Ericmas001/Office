using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ExcelConditionnalFormating
    {
        public XlFormatConditionType Type { get; set; }
        public XlFormatConditionOperator Operator { get; set; }
        public string Formula { get; set; }
        public ExcelFormat Format { get; set; }
    }
}
