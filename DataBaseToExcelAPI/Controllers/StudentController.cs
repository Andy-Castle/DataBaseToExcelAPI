using ClosedXML.Excel;
using DataBaseToExcelAPI.Data;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DataBaseToExcelAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {

        private readonly ApplicationDbContext _context;

        public StudentController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet("ExportToExcel")]
        public IActionResult ExportToExcel()
        {
            try
            {

                var students = _context.Students.ToList();

                if (students.Count == 0)
                {
                    return StatusCode(404, "No hay información en la base de datos");

                }

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Students");

                        //Se agregan los encabezado
                        worksheet.Cell(1, 1).Value = "ID";
                        worksheet.Cell(1, 2).Value = "Name";
                        worksheet.Cell(1, 3).Value = "Marks";


                        //Se da formato a los encabezados
                        var headerRange = worksheet.Range(1,1,1,3);
                        headerRange.Style.Font.SetBold()
                            .Fill.SetBackgroundColor(XLColor.AliceBlue)
                            .Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                        //Se agregan los datos
                        int row = 2;

                        foreach (var student in students)
                        {
                            worksheet.Cell(row, 1).Value = student.Id;
                            worksheet.Cell(row, 2).Value = student.Name;
                            worksheet.Cell(row, 3).Value = student.Marks;
                            row++;
                        }

                        //Se ajustan las columnas
                        worksheet.Columns().AdjustToContents();

                        //Se convierte en memorystream
                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);

                            return File(
                                stream.ToArray(),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                $"Students_{DateTime.Now:yyyyMMdd}.xlsx"
                                );
                        }
                    }
                

            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    }
}
