using System;
using System.Collections.Generic;
using System.Linq;
using FastExcel;

namespace templater
{
    public class Utils
    {
        public static int GetColumnNumberByName(string name, Worksheet worksheet)
        {
            int column = -1;
            foreach (var titleCell in worksheet.Rows.FirstOrDefault().Cells)
            {
                if ((string)titleCell.Value == name)
                {
                    column = titleCell.ColumnNumber;
                }
            }

            if (column == -1)
            {
                throw new Exception($"Nie znaleziono kolumny zatytułowanej {name}!");
            }

            return column;
        }

        public static List<T> ReadAllButFirstCellsInColumn<T>(int colNum, Worksheet worksheet)
        {
            List<T> list = new();
            
            foreach (var tekstCellRow in worksheet.Rows.ToArray()[1..])
            {
                list.Add((T)(tekstCellRow.Cells.FirstOrDefault(c => c.ColumnNumber == colNum).Value ?? ""));
            }

            return list;
        }

        public static List<T> GetCellsAndNumber<T>(string name, Worksheet worksheet)
        {
            List<T> list = new();
            var columnNum = Utils.GetColumnNumberByName(name, worksheet);
            return Utils.ReadAllButFirstCellsInColumn<T>(columnNum, worksheet);
        }
    }
}