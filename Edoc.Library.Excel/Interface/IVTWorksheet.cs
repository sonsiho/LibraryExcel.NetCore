using Aspose.Cells;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorksheet
    {
        Worksheet Worksheet { get; set; }
        string Name { get; set; }

        IVTWorksheet SetCellValue(string cell, object value);

        IVTWorksheet SetCellValue(int row, int column, object value);

        object GetCellValue(int row, int column);

        object GetCellValue(string cell);

        T GetCellValue<T>(string cell);

        T GetCellValue<T>(int row, int column);
        IVTWorksheet CopyAndInsertARow(int copiedRow, int insertRow);
    }
}
