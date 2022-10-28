using Edoc.Library.Excel.Common;
using Edoc.Library.Excel.Interface;
using NativeExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    internal class VTWorksheet : IVTWorksheet
    {
        private readonly IWorksheet _worksheet;
        private IVTRange _vtCells { get; set; }
        public VTWorksheet(IWorksheet worksheet)
        {
            this._worksheet = worksheet;

            _vtCells = new VTRange(worksheet.Range);
        }

        public string Name { get => this._worksheet.Name; set => this._worksheet.Name = value; }

        public int Index => this._worksheet.Index;

        public IVTRange Cells => this._vtCells;

        public object GetCellValue(int row, int column)
        {
            return this._worksheet.Cells[row, column].Value;
        }

        public object GetCellValue(string cell)
        {
            return this._worksheet.Cells[cell].Value;
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

        public void SetCellValue(string cell, object value)
        {
            NativeExcel.IRange iCell = _worksheet.Range[cell];
            iCell.Formula = value;
        }

        public void SetCellValue(int row, int column, object value)
        {
            NativeExcel.IRange iCell = _worksheet.Range[row, column];
            iCell.Formula = value;
        }
    }
}
