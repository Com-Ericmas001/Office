using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ExcelBorder
    {
        public XlLineStyle BorderStyle { get; set; }
        public Color? BorderColor { get; set; }
        public double? BorderThickness { get; set; }
    }
}
