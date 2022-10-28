using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorksheets : IEnumerable
    {
        IVTWorksheet this[int index] { get; }

        IVTWorksheet this[string Name] { get; }

        int Count { get; }
    }
}
