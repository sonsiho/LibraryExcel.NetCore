using Edoc.Library.Excel.Factory;
using Edoc.Library.Excel.Interface;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    public class VTWorkbook : VTWorkbookFactory, IVTWorkbook
    {
        private NativeExcel.IWorkbook _workbook { get; set; }

        public VTWorkbook(string templateFilePath) : base(templateFilePath)
        {
            _workbook = NativeExcel.Factory.OpenWorkbook(templateFilePath);
        }

        public async Task<byte[]> ToByteArrAsync()
        {
            _workbook.SaveAs(TempFilePath);
            this.RemoveLicense();
            var result = await File.ReadAllBytesAsync(TempFilePath);
            File.Delete(TempFilePath);
            return result;
        }

        public void SaveAs(string filePath)
        {
            _workbook.SaveAs(TempFilePath);
            this.RemoveLicense();
            File.Copy(TempFilePath, filePath);
        }

        public void Protect()
        {
            throw new NotImplementedException();
        }

        public void Protect(string Password)
        {
            throw new NotImplementedException();
        }

        public void Unprotect()
        {
            throw new NotImplementedException();
        }

        public void Unprotect(string Password)
        {
            throw new NotImplementedException();
        }
    }
}
