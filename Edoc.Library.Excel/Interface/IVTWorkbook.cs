using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Interface
{
    public interface IVTWorkbook
    {
        /// <summary>
        /// Chuyển sang định dạng Stream
        /// </summary>
        Task<byte[]> ToByteArrAsync();

        /// <summary>
        /// Lưu vào một file
        /// </summary>
        /// <param name="pathFile">Đường dẫn file + tên file</param>
        void SaveAs(string pathFile);

        void Protect();

        void Protect(string Password);

        void Unprotect();

        void Unprotect(string Password);
    }
}
