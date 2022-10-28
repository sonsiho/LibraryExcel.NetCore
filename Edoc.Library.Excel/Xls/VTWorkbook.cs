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
        private IVTWorksheets _vtWorkSheets { get; set; }

        public bool IsProtected => this._workbook.IsProtected;

        public IVTWorksheets Worksheets => this._vtWorkSheets;

        public VTWorkbook(string templateFilePath) : base(templateFilePath)
        {
            _workbook = NativeExcel.Factory.OpenWorkbook(templateFilePath);
            _vtWorkSheets = new VTWorksheets(_workbook.Worksheets);
           
        }

        public VTWorkbook(string templateFilePath,string password) : base(templateFilePath, password)
        {
            _workbook = NativeExcel.Factory.OpenWorkbook(templateFilePath, password);
            _vtWorkSheets = new VTWorksheets(_workbook.Worksheets);
        }

        public async Task<byte[]> ToByteArrAsync()
        {
            _workbook.SaveAs(TempFilePath);
            this.RemoveLicense();
            var result = await File.ReadAllBytesAsync(TempFilePath);
            File.Delete(TempFilePath);
            return result;
        }

        public async Task SaveAsAsync(string filePath)
        {
            var byteArray = await this.ToByteArrAsync();

            await File.WriteAllBytesAsync(filePath, byteArray);
        }

        public IVTWorkbook Protect(string Password)
        {
            this._workbook.Protect(Password);

            return this;
        }

        public IVTWorkbook Unprotect(string Password)
        {
            this._workbook.Unprotect(Password);

            return this;
        }
    }
}
