using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Com.Ericmas001.Office.Excel.ExportParms
{
    public class ExcelExportParm
    {
        private Dictionary<string, ColumnExportParm> m_ColumnParms = new Dictionary<string, ColumnExportParm>();
        public bool UseView { get; set; }
        public ExcelFormat HeaderFormat{ get; set; }

        public ExcelFormat DefaultCellFormat { get; set; }

        public Dictionary<string, ColumnExportParm> ColumnParms
        {
            get { return m_ColumnParms; }
            set { m_ColumnParms = value; }
        }

        public ExcelExportParm()
        {
            UseView = false;
            DefaultCellFormat = new ExcelFormat
            {
                Border = new ExcelBorder
                {
                    BorderStyle = XlLineStyle.xlContinuous,
                    BorderColor = Color.DarkGray,
                    BorderThickness = 1
                }
            };
                
            HeaderFormat = new ExcelFormat
            {
                ForeColor = Color.White,
                BackColor = Color.Black,
                Bold = true
            };
        }
    }
}
