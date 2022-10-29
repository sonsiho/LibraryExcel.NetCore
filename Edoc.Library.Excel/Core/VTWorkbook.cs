using Aspose.Cells;
using Edoc.Library.Excel.Common;
using Edoc.Library.Excel.Interface;
using System;
using System.Collections.Generic;
using System.IO;

namespace Edoc.Library.Excel.Core
{
    internal class VTWorkbook : IVTWorkbook
    {
        public Workbook Workbook { get; set; }

        public VTWorkbook(VtFileFormatType fileFormatType)
        {
            this.SetLicense();

            this.Workbook = new Workbook((FileFormatType)fileFormatType);
        }

        public VTWorkbook(string filePath, string password = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' cannot be null or whitespace.", nameof(filePath));
            }

            this.SetLicense();

            var loadOptions = this.BuildLoadOptions(password);

            this.Workbook = new Workbook(filePath, loadOptions);
        }

        public VTWorkbook(Stream stream, string password = null)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            this.SetLicense();

            var loadOptions = this.BuildLoadOptions(password);

            this.Workbook = new Workbook(stream, loadOptions);
        }

        public void Save(string filePath, VtSaveFormat saveFormat = VtSaveFormat.Auto)
        {
            Workbook.Save(filePath, (SaveFormat)saveFormat);
        }

        public Stream ToStream(VtSaveFormat saveFormat = VtSaveFormat.Auto)
        {
            var tempFile = Path.GetTempFileName();
            using FileStream fileStream = new FileStream(tempFile, FileMode.CreateNew);
            Workbook.Save(fileStream, (SaveFormat)saveFormat);

            var memStream = new MemoryStream();
            memStream.SetLength(fileStream.Length);
            fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
            return memStream;
        }

        public Stream ToStream()
        {
            return Workbook.SaveToStream();
        }

        public IVTWorksheet GetSheet(int index)
        {
            return new VTWorksheet(Workbook.Worksheets[index]);
        }

        public IVTWorksheet GetSheet(string sheetName)
        {
            return new VTWorksheet(Workbook.Worksheets[sheetName]);
        }

        public List<IVTWorksheet> GetSheets()
        {
            List<IVTWorksheet> list = new List<IVTWorksheet>();
            foreach (var sheet in Workbook.Worksheets)
                list.Add(new VTWorksheet(sheet));
            return list;
        }

        private void SetLicense()
        {
            License license = new License();

            license.SetLicense("lib\\Aspose\\Aspose.License.xml");
        }

        private LoadOptions BuildLoadOptions(string password = null)
        {
            var loadOptions = new LoadOptions();
            if (!string.IsNullOrWhiteSpace(password))
            {
                loadOptions.Password = password;
            }

            return loadOptions;
        }        
    }
}
