using Microsoft.AspNetCore.Http;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace CheckExcel.Services
{
    public class CustomerDetail : ICustomer
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _environment;
        public CustomerDetail(IConfiguration configuration, IWebHostEnvironment environment)
        {
            _configuration = configuration;
            _environment = environment;
        }
        public DataTable CustomerDataTable(string path)
        {
            var constr = _configuration.GetConnectionString("excelconnection");
            DataTable datatable = new DataTable();

            constr = string.Format(constr, path);

            using (OleDbConnection excelconn = new OleDbConnection(constr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter adapterexcel = new OleDbDataAdapter())
                    {

                        excelconn.Open();
                        cmd.Connection = excelconn;
                        DataTable excelschema;
                        excelschema = excelconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        var sheetname = excelschema.Rows[0]["Table_Name"].ToString();
                        excelconn.Close();

                        excelconn.Open();
                        cmd.CommandText = "SELECT * From [" + sheetname + "]";
                        adapterexcel.SelectCommand = cmd;
                        adapterexcel.Fill(datatable);
                        excelconn.Close();

                    }
                }

            }

            return datatable;
        }

        public string Documentupload(IFormFile fromFiles)
        {
            string uploadpath = _environment.WebRootPath;
            string dest_path = Path.Combine(uploadpath, "uploaded_doc");

            if (!Directory.Exists(dest_path))
            {
                Directory.CreateDirectory(dest_path);
            }
            string sourcefile = Path.GetFileName(fromFiles.FileName);
            string path = Path.Combine(dest_path, sourcefile);

            using (FileStream filestream = new FileStream(path, FileMode.Create))
            {
                fromFiles.CopyTo(filestream);
            }
            return path;
        }

        public void ImportCustomer(DataTable customer)
        {

            var sqlconn = _configuration.GetConnectionString("sqlconnection");
            using (SqlConnection scon = new SqlConnection(sqlconn))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(scon))
                {
                    sqlBulkCopy.DestinationTableName = "Customers";


                    sqlBulkCopy.ColumnMappings.Add("FirstName", "firstName");
                    sqlBulkCopy.ColumnMappings.Add("LastName", "lastName");
                    sqlBulkCopy.ColumnMappings.Add("job", "job");
                    sqlBulkCopy.ColumnMappings.Add("amount", "amount");
                    sqlBulkCopy.ColumnMappings.Add("hiredate", "tdate");

                    scon.Open();
                    sqlBulkCopy.WriteToServer(customer);
                    scon.Close();
                }

            }
        }
    }
}