using Aspose.Cells;
using Edoc.Library.Excel.Common;
using Edoc.Library.Excel.Interface;
using GemBox.Spreadsheet;

namespace Edoc.Library.Excel.Core
{
    internal class VTWorksheet : IVTWorksheet
    {
        public Worksheet Worksheet { get; set; }
        public VTWorksheet(Worksheet worksheet)
        {
            this.Worksheet = worksheet;
        }

        public string Name { get => this.Worksheet.Name; set => this.Worksheet.Name = value; }

        public object GetCellValue(int row, int column)
        {
            return this.Worksheet.Cells[row - 1, column - 1].Value;
        }

        public object GetCellValue(string cell)
        {
            return this.Worksheet.Cells[cell].Value;
        }

        public T GetCellValue<T>(string cell)
        {
            var Value = this.GetCellValue(cell);
            if (Value == null)
                return default(T);
            return Value.ConvertToOrDefault<T>();
        }

        public T GetCellValue<T>(int row, int column)
        {
            var Value = this.GetCellValue(row, column);
            if (Value == null)
                return default(T);
            return Value.ConvertToOrDefault<T>();
        }

        public IVTWorksheet SetCellValue(string cell, object value)
        {
            Worksheet.Cells[cell].Value = value;

            return this;
        }

        public IVTWorksheet SetCellValue(int row, int column, object value)
        {
            Worksheet.Cells[row - 1, column - 1].Value = value;

            return this;
        }

        public IVTWorksheet CopyAndInsertARow(int copiedRow, int insertRow)
        {
            this.Worksheet.Cells.CopyRow(Worksheet.Cells, copiedRow - 1, insertRow - 1);

            return this;
        }

        public IVTWorksheet CopyAndInsertARow(int copiedRow, int insertRow)
        {
            this.Worksheet.Cells.CopyRow(Worksheet.Cells, copiedRow - 1, insertRow - 1);

            return this;
        }
    }
}
