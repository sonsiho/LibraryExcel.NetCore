using Edoc.Library.Excel;
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

            var workBook = EdocExcel.OpenWorkBook(template);

            var sheet = workBook.GetSheet(0);

            sheet.CopyAndInsertARow(2,14);
            workBook.Save(result);

            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}
