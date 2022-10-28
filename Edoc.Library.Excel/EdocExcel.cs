using Edoc.Library.Excel.Interface;
using Edoc.Library.Excel.Xls;
using System;
using System.IO;

namespace Edoc.Library.Excel
{
    public static class EdocExcel
    {
        public static IVTWorkbook OpenWorkBook(string templateFilePath)
        {
            var fileExtension = Path.GetExtension(templateFilePath);
            if (fileExtension == ".xls")
            {
                return new VTWorkbook(templateFilePath);
            }
            else
            {
                throw new Exception("Invalid excel path");
            }
        }
    }
}
