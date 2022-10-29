using Aspose.Cells;
using Edoc.Library.Excel.Common;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorkbook
    {
        Workbook Workbook { get; set; }

        void Save(string filePath, VtSaveFormat saveFormat = VtSaveFormat.Auto);

        Stream ToStream();

        Stream ToStream(VtSaveFormat saveFormat = VtSaveFormat.Auto);

        List<IVTWorksheet> GetSheets();

        IVTWorksheet GetSheet(int index);

        IVTWorksheet GetSheet(string sheetName);
    }
}
