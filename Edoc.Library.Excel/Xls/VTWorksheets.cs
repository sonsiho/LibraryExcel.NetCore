using Edoc.Library.Excel.Interface;
using NativeExcel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    internal class VTWorksheets : IVTWorksheets
    {
        private List<IVTWorksheet> _vtWorksheets = new List<IVTWorksheet>();
        public VTWorksheets(IWorksheets workSheets) 
        {
            foreach (IWorksheet item in workSheets)
            {
                _vtWorksheets.Add(new VTWorksheet(item));
            }
            
        }

        public IVTWorksheet this[int index] => _vtWorksheets[index];

        public IVTWorksheet this[string Name]
        {
            get
            {
                var vtWorkSheet = _vtWorksheets.FirstOrDefault(n=>n.Name.ToLower() == Name.ToLower());
                return vtWorkSheet;
            }
        }

        public int Count => this._vtWorksheets.Count;

        public IEnumerator GetEnumerator()
        {
            return this._vtWorksheets.GetEnumerator();
        }
    }
}
