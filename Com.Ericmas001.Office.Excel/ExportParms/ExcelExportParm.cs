using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ExcelExportParm
    {
        private Dictionary<string, ColumnExportParm> m_ColumnParms = new Dictionary<string, ColumnExportParm>();
        public bool UseView { get; set; }
        public ExcelFormat HeaderFormat{ get; set; }

        public Dictionary<string, ColumnExportParm> ColumnParms
        {
            get { return m_ColumnParms; }
            set { m_ColumnParms = value; }
        }

        public ExcelExportParm()
        {
            UseView = false;
            HeaderFormat = new ExcelFormat
            {
                ForeColor = Color.White,
                BackColor = Color.Black,
                Bold = true
            };
        }
    }
}
