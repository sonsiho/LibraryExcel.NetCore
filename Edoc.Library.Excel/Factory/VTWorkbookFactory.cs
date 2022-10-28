using System;
using System.IO;
using System.Threading.Tasks;
using Edoc.Library.Excel.Interface;
using GemBox.Spreadsheet;

namespace Edoc.Library.Excel.Factory
{
    public class VTWorkbookFactory
    {
        public VTWorkbookFactory(string templateFilePath)
        {
            if (string.IsNullOrWhiteSpace(templateFilePath))
            {
                throw new ArgumentException($"'{nameof(templateFilePath)}' cannot be null or whitespace.", nameof(templateFilePath));
            }

            TemplateFilePath = templateFilePath;

            FileName = Path.GetFileName(TemplateFilePath);

            FileNameWithoutExtension = Path.GetFileNameWithoutExtension(TemplateFilePath);

            Extension = Path.GetExtension(TemplateFilePath);

            TempFilePath = Path.Combine(Path.GetTempPath(),$"{Guid.NewGuid()}{this.Extension}");
        }

        public string TemplateFilePath { get; set; }

        public string FileName { get; set; }

        public string Extension { get; set; }

        public string FileNameWithoutExtension { get; set; }

        public string TempFilePath { get; set; }

        protected void RemoveLicense()
        {
            ExcelFile workbook = ExcelFile.Load(TempFilePath);
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Rows.Remove(0);
            }

            workbook.Save(TempFilePath);
        }
    }
}
