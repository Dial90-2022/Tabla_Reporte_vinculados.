using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Tabla_Reporte_vinculados_.Models;
using System.Data;
using System.Data.SqlClient;

using ClosedXML.Excel;

namespace Tabla_Reporte_vinculados_.Controllers
{
    public class HomeController : Controller
    {
        private readonly string cadenaSQL;

        public HomeController(IConfiguration config)
        {
            cadenaSQL = config.GetConnectionString("cadenaSQL");
        }

        public IActionResult Index()
        {
            return View();
        }


        public IActionResult Exportar_Excel(string fechaInicio, string fechaFin)
        {
            DataTable tabla_Vinculados = new DataTable();

            //=========== PRIMERO - OBTENER EL DATA ADAPTER ===========
            using (var conexion = new SqlConnection(cadenaSQL))
            {
                conexion.Open();
                using (var adapter = new SqlDataAdapter())
                {

                    adapter.SelectCommand = new SqlCommand("sp_Reporte_Vinculados", conexion);
                    adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaInicio", fechaInicio);
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaFin", fechaFin);



                    adapter.Fill(tabla_Vinculados);
                }
            }

            //usar referencias
            //=========== Se usa  ClosedXML para exportar el excel===========
            using (var libro = new XLWorkbook())
            {

                tabla_Vinculados.TableName = "Vinculados";
                var hoja = libro.Worksheets.Add(tabla_Vinculados);
                hoja.ColumnsUsed().AdjustToContents();

                using (var memoria = new MemoryStream())
                {

                    libro.SaveAs(memoria);

                    var nombreExcel = string.Concat("Reporte ", DateTime.Now.ToString(), ".xlsx");

                    return File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);






                }
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}