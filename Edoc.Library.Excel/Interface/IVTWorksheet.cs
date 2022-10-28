using NativeExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorksheet
    {
        string Name { get; set; }

        int Index { get; }

        IVTRange Cells { get; }

        void SetCellValue(string cell, object value);

        void SetCellValue(int row, int column, object value);

        object GetCellValue(int row, int column);

        object GetCellValue(string cell);

        T GetCellValue<T>(string cell);

        T GetCellValue<T>(int row, int column);
    }
}
