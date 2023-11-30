using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace Exceldatareader.Controllers
{
        public class HomeController : Controller
        {
            private string connectionString = "Server=FGLAPNL207HFZT\\SQLEXPRESS;Database=uploadingfile;Trusted_Connection=True;MultipleActiveResultSets=true; TrustServerCertificate=True ";

            public ActionResult Index()
            {
                return View();
            }

            [HttpPost]
            public ActionResult Upload(HttpPostedFileBase file)
            {
                if (file != null && file.ContentLength > 0)
                {
                    try
                    {
                        string fileName = Path.GetFileName(file.FileName);
                        string filePath = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                        file.SaveAs(filePath);

                        DataTable dt = ReadExcel(filePath);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            CreateAndSaveTable(dt, "NewTable");
                            ViewBag.Message = "Table created successfully!";
                        }
                        else
                        {
                            ViewBag.Message = "No data found in the Excel file.";
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = "Error: " + ex.Message;
                    }
                }
                else
                {
                    ViewBag.Message = "Please select a file to upload.";
                }

                return View("Index");
            }

            private DataTable ReadExcel(string filePath)
            {
                using (OleDbConnection conn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;'"))
                {
                    conn.Open();
                    DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = dt?.Rows?[0]?["TABLE_NAME"]?.ToString();

                    if (sheetName != null)
                    {
                        OleDbCommand cmd = new OleDbCommand($"SELECT * FROM [{sheetName}]", conn);
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        DataTable excelDt = new DataTable();
                        da.Fill(excelDt);
                        return excelDt;
                    }
                }
                return null;
            }

            private void CreateAndSaveTable(DataTable dt, string tableName)
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    SqlCommand createCommand = new SqlCommand($"CREATE TABLE {tableName} (", connection);
                    foreach (DataColumn column in dt.Columns)
                    {
                        createCommand.CommandText += $"{column.ColumnName} NVARCHAR(255), ";
                    }
                    createCommand.CommandText = createCommand.CommandText.TrimEnd(',', ' ') + ")";
                    createCommand.ExecuteNonQuery();

                    foreach (DataRow row in dt.Rows)
                    {
                        SqlCommand insertCommand = new SqlCommand($"INSERT INTO {tableName} VALUES (", connection);
                        foreach (var item in row.ItemArray)
                        {
                            insertCommand.CommandText += $"'{item.ToString()}', ";
                        }
                        insertCommand.CommandText = insertCommand.CommandText.TrimEnd(',', ' ') + ")";
                        insertCommand.ExecuteNonQuery();
                    }
                }
            }
        }
    }



