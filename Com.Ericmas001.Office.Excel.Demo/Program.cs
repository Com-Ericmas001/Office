using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Com.Ericmas001.Office.Excel.ExportParms;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Com.Ericmas001.Office.Excel.Demo
{
    class Program
    {
        static void Main()
        {
            const string ID = "ID";
            const string FRUIT = "Fruit";
            const string FAVORITE = "Favorite";

            var table = new DataTable();
            table.Columns.Add(ID);
            table.Columns.Add(FRUIT);
            table.Columns.Add(FAVORITE);

            string[] fruits = {"Apple", "Orange", "Pineapple", "Grapes", "Peach", "Pear", "Strawberry", "Raspberry", "BlueBerry", "Lemon"};
            for (var i = 0; i < 100; i++)
            {
                var row = table.NewRow();
                row[ID] = (i + 1).ToString();
                row[FRUIT] = fruits[i%fruits.Length];
                row[FAVORITE] = (i*(i + 2))%7 == 0 ? "X" : "";
                table.Rows.Add(row);
            }

            var parms = new ExcelExportParm();

            var favColPArms = new ColumnExportParm
            {
                HorizontalAlignment = XlHAlign.xlHAlignCenter,
                FormatConditions = new []
                {
                    new ExcelConditionnalFormating
                    {
                        Type = XlFormatConditionType.xlCellValue,
                        Operator = XlFormatConditionOperator.xlEqual,
                        Formula = "X",
                        Format = new ExcelFormat
                        {
                            ForeColor = Color.Red,
                            BackColor = Color.Bisque,
                            Bold = true
                        }
                    }
                }
            };

            parms.ColumnParms.Add(FAVORITE, favColPArms);

            new ExcelExporter().ExportDataTable(table, parms);
        }
    }
}
