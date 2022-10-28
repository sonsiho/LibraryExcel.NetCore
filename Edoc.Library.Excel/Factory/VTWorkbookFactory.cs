using System;
using System.IO;
using System.Text;
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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            TemplateFilePath = templateFilePath;

            this.Init();
        }

        public VTWorkbookFactory(string templateFilePath,string password)
        {
            if (string.IsNullOrWhiteSpace(templateFilePath))
            {
                throw new ArgumentException($"'{nameof(templateFilePath)}' cannot be null or whitespace.", nameof(templateFilePath));
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException($"'{nameof(password)}' cannot be null or whitespace.", nameof(password));
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            TemplateFilePath = templateFilePath;
            PassWord = password;
            this.Init();
        }

        private void Init()
        {
            FileName = Path.GetFileName(TemplateFilePath);

            FileNameWithoutExtension = Path.GetFileNameWithoutExtension(TemplateFilePath);

            Extension = Path.GetExtension(TemplateFilePath);

            TempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}{this.Extension}");
        }

        public string TemplateFilePath { get; set; }

        public string FileName { get; set; }

        public string Extension { get; set; }

        public string FileNameWithoutExtension { get; set; }

        public string TempFilePath { get; set; }

        public string PassWord { get; set; }

        protected void RemoveLicense()
        {
            ExcelFile workbook = ExcelFile.Load(TempFilePath);
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Rows[0].Cells[0].Value = string.Empty;
            }

            workbook.Save(TempFilePath);
        }
    }
}
