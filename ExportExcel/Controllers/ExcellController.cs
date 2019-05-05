using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExportExcel.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExportExcel.Controllers
{
    [Route("api/excel")]
    public class ExcellController : Controller
    {
        [HttpPost]
         public async Task<ResultObject<List<Employees>>> Import( IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return ResultObject<List<Employees>>.GetResult(-1, "formfile is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return ResultObject<List<Employees>>.GetResult(-1, "Not Support file extension");
            }

            var list = new List<Employees>();

            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var a = worksheet;
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        list.Add(new Employees
                        {
                            FirstName = worksheet.Cells[row, 1].Value.ToString().Trim(),
                            LastName = worksheet.Cells[row, 2].Value.ToString().Trim(),
                            Gender = worksheet.Cells[row, 3].Value.ToString().Trim(),
                            Salary = int.Parse(worksheet.Cells[row, 4].Value.ToString().Trim()),
                        });
                    }
                }
            }

            // add list to db ..  
            // here just read and return  

            return ResultObject<List<Employees>>.GetResult(200, "OK", list);
        }
    }
}