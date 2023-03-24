using System.Data;

namespace CheckExcel.Services
{
    public interface ICustomer
    {
        string Documentupload(IFormFile formFile);
        DataTable CustomerDataTable(string path);
        void ImportCustomer(DataTable customer);
    }
}
