using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CustomExtensions
{
    public static class DGVExtentions
    {
        public static DataGridViewRow CloneWithValues(this DataGridViewRow row)
        {
            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
            for (Int32 index = 0; index < row.Cells.Count; index++)
            {
                clonedRow.Cells[index].Value = row.Cells[index].Value;
            }
            return clonedRow;
        } // End CloneWithValues()

        public static void CloneColumns(this DataGridView Self, DataGridView CloneFrom)
        {
            if (Self.ColumnCount == 0)
            {
                foreach (DataGridViewColumn c in CloneFrom.Columns)
                {
                    Self.Columns.Add((DataGridViewColumn)c.Clone());
                }
            }
        } // End CloneColumns()
    }

    public static class ExcelExtensions
    {
        /// <summary>
        /// Returns the value in the selected cell.
        /// </summary>
        /// <param name="sheet">The sheet the cells are located on.</param>
        /// <param name="x">Row number of the cell.</param>
        /// <param name="y">Column number of the cell.</param>
        /// <returns>Value in the selected cell.</returns>
        public static string GetCellValue(this Worksheet sheet, int x, int y)
        {            
            string value = (((Range)sheet.Cells[x, y]).Value2);
            return value;
        }

        /// <summary>
        /// Changes the value of an Excel Sheet cell based on type.
        /// </summary>
        /// <param name="sheet">The sheet the cells are located on.</param>
        /// <param name="x">Row number of the cell.</param>
        /// <param name="y">Column number of the cell.</param>
        /// <param name="type">Type of the new value.</param>
        /// <param name="value">What is put into the cell.</param>
        public static void SetCellValue(this Worksheet sheet, int x, int y, string type, string value)
        {
            try
            {
                switch (type.ToLower())
                {
                    case "string":
                        (((Range)sheet.Cells[x, y]).Value2) = value;
                        break;
                    case "long":
                        (((Range)sheet.Cells[x, y]).Value2) = long.Parse(value);
                        break;
                    case "int":
                        (((Range)sheet.Cells[x, y]).Value2) = int.Parse(value);
                        break;
                    case "bool":
                        (((Range)sheet.Cells[x, y]).Value2) = bool.Parse(value);
                        break;
                    case "float":
                        (((Range)sheet.Cells[x, y]).Value2) = float.Parse(value);
                        break;
                    case "decimal":
                        (((Range)sheet.Cells[x, y]).Value2) = bool.Parse(value);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not convert " + value + " to type " + type + ".  Error is as follows: \n\n"+ex.Message,"Error in Conversion",MessageBoxButtons.OK,MessageBoxIcon.Error );
            }
        }
    }
}