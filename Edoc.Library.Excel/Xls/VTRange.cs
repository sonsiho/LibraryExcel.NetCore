using Edoc.Library.Excel.Interface;
using NativeExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    internal class VTRange : IVTRange
    {
        private readonly IRange _range;

        public VTRange(IRange range)
        {
            this._range = range;
        }
    }
}
