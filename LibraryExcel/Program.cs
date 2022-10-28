using Edoc.Library.Excel;
using GemBox.Spreadsheet;
using System;
using System.Threading.Tasks;

namespace LibraryExcel
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var template = @"C:\Data\REPOSITORIES\LibraryExcel\LibraryExcel\Template\Edoc_Test.xls";
            var result = @"C:\Data\REPOSITORIES\LibraryExcel\LibraryExcel\Template\Edoc_Test_result.xls";

            ExcelFile workbook = ExcelFile.Load(template);
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Rows[0].Cells[0].Value = "1";
            }

            workbook.Save(result);

            var workBook = EdocExcel.OpenWorkBook(template);

            var sheet = workBook.Worksheets[0];

            var test = sheet.GetCellValue("B11");
            sheet.SetCellValue("B11", "Test");
            await workBook.SaveAsAsync(result);
            Console.ReadLine();
        }
    }
}
