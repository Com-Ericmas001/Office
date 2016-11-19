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
            const string DESC = "Description";

            var table = new DataTable();
            table.Columns.Add(ID);
            table.Columns.Add(FRUIT);
            table.Columns.Add(FAVORITE);
            table.Columns.Add(DESC);

            string[] fruits = {"Apple", "Orange", "Pineapple", "Grapes", "Peach", "Pear", "Strawberry", "Raspberry", "BlueBerry", "Lemon"};
            for (var i = 0; i < 100; i++)
            {
                var row = table.NewRow();
                row[ID] = (i + 1).ToString();
                row[FRUIT] = fruits[i%fruits.Length];
                row[FAVORITE] = (i * (i + 2)) % 7 == 0 ? "X" : "";
                row[DESC] = fruits[i % fruits.Length] + Environment.NewLine + fruits[i % (fruits.Length/2)];
                table.Rows.Add(row);
            }

            var parms = new ExcelExportParm
            {
                HeaderFormat = new ExcelFormat
                {
                    Bold = true,
                    BackColor = Color.Aqua,
                    ForeColor = Color.Black,
                    Border = new ExcelBorder()
                    {
                        BorderStyle = XlLineStyle.xlContinuous,
                        BorderThickness = 4
                    }
                },
                DefaultCellFormat = new ExcelFormat()
                {
                    ForeColor = Color.BlueViolet,
                    Border = new ExcelBorder()
                    {
                        BorderColor = Color.Red,
                        BorderStyle = XlLineStyle.xlContinuous,
                        BorderThickness = 2d
                    }
                }
            };

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
            var fruitColPArms = new ColumnExportParm
            {
                EnumValues = fruits
            };

            parms.ColumnParms.Add(FRUIT, fruitColPArms);

            new ExcelExporter().ExportDataTable(table, parms);
        }
    }
}
