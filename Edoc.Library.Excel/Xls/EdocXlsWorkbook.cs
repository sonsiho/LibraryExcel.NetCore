using Edoc.Library.Excel.Factory;
using Edoc.Library.Excel.Interface;
using NativeExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Xls
{
    public class EdocXlsWorkbook : EdocWorkbookFactory, IEdocWorkbook
    {
        public NativeExcel.IWorkbook Workbook { get; set; }

        public EdocXlsWorkbook(string templateFilePath) : base(templateFilePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Workbook = NativeExcel.Factory.OpenWorkbook(templateFilePath);
        }

        public IEnumerable<IEdocWorksheet> GetSheets()
        {
            List<IEdocWorksheet> list = new List<IEdocWorksheet>();
            foreach (NativeExcel.IWorksheet sheet in Workbook.Worksheets)
            {
                list.Add(new EdocXlsWorksheet(sheet));
            }
                
            return list;
        }

        public IEdocWorksheet GetSheet(int index)
        {
            return new EdocXlsWorksheet(Workbook.Worksheets[index]);
        }

        public IEdocWorksheet GetSheet(string sheetName)
        {
            return new EdocXlsWorksheet(Workbook.Worksheets[sheetName]);
        }

        public async Task<byte[]> ToByteArrayAsync()
        {
            Workbook.SaveAs(TempFilePath);
            this.RemoveLicense();
            var result = await File.ReadAllBytesAsync(TempFilePath);
            File.Delete(TempFilePath);
            return result;
        }

        public async Task SaveAsAsync(string filePath)
        {
            Workbook.SaveAs(TempFilePath);
            this.RemoveLicense();
            var result = await File.ReadAllBytesAsync(TempFilePath);
            await File.WriteAllBytesAsync(filePath, result);
            File.Delete(TempFilePath);
        }

        public IEdocWorksheet CopySheetToLast(IEdocWorksheet worksheet, string lastCell = null)
        {
            NativeExcel.IWorksheet templateSheet = worksheet.Worksheet;
            NativeExcel.IWorksheet sheet = Workbook.Worksheets.AddAfter(Workbook.Worksheets.Count);

            IVTWorksheet vtSheet = new VTWorksheet(sheet);
            if (lastCell == null)
            {
                vtSheet.CopySheet(worksheet);
            }
            else
            {
                vtSheet.CopySheet(worksheet, lastCell);
            }
            return vtSheet;
        }
    }
}
