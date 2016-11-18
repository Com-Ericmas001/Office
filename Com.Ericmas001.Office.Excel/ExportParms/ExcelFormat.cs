using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ExcelFormat
    {
        public Color ForeColor { get; set; }
        public Color BackColor { get; set; }
        public bool Bold { get; set; }
    }
}
