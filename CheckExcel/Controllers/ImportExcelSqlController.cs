using CheckExcel.Services;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace CheckExcel.Controllers
{
    public class ImportExcelSqlController : Controller
    {
        private IWebHostEnvironment _environment;
        private readonly ICustomer _customer;
        //  private IConfiguration _configuration;
        public ImportExcelSqlController(IWebHostEnvironment environment, ICustomer customer)
        {
            _environment = environment;
            _customer = customer;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [AutoValidateAntiforgeryToken]
        public IActionResult Index(IFormFile fromFiles)
        {
            string path = _customer.Documentupload(fromFiles);
            DataTable dt = _customer.CustomerDataTable(path);
            _customer.ImportCustomer(dt);
            return View();
        }
    }
}