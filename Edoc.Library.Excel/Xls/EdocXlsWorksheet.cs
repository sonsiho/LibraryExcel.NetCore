using Edoc.Library.Excel.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    public class EdocXlsWorksheet : IEdocWorksheet
    {
        public NativeExcel.IWorksheet Worksheet { get; set; }

        public EdocXlsWorksheet(NativeExcel.IWorksheet worksheet)
        {
            Worksheet = worksheet;
        }

    }
}
