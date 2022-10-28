using System.Threading.Tasks;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorkbook
    {
        IVTWorksheets Worksheets { get; }

        Task<byte[]> ToByteArrAsync();

        Task SaveAsAsync(string pathFile);


        IVTWorkbook Protect(string Password);

        IVTWorkbook Unprotect(string Password);

        bool IsProtected { get; }
    }
}
