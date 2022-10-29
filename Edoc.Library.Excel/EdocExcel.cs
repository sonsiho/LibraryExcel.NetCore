using Edoc.Library.Excel.Common;
using Edoc.Library.Excel.Core;
using Edoc.Library.Excel.Interface;
using System;
using System.IO;

namespace Edoc.Library.Excel
{
    public static class EdocExcel
    {
        public static IVTWorkbook OpenWorkBook(string templateFilePath, string password = null)
        {
            return new VTWorkbook(templateFilePath, password);
        }

        public static IVTWorkbook CreateWorkBook(VtFileFormatType fileFormatType)
        {
            return new VTWorkbook(fileFormatType);
        }
    }
}
