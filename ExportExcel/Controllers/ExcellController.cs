using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExportExcel.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExportExcel.Controllers
{
    [Route("api/excel")]
    public class ExcellController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public ExcellController(IHostingEnvironment host)
        {
            this._hostingEnvironment = host;
        }
        [HttpPost]
        public async Task<ResultObject<List<Employees>>> Import(IFormFile formFile, CancellationToken cancellationToken)
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

        [HttpGet("export")]
        public async Task<ResultObject<string>> Export(CancellationToken cancellationToken)
        {
            string folder = "File";
            string excelName = $"ShifLog-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, excelName);
            FileInfo file = new FileInfo(Path.Combine(folder, excelName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(folder, excelName));
            }

            // query data from database  
            await Task.Yield();

            var list = new List<Employees>()
    {
        new Employees { FirstName = "Tuan",LastName="Anh",Gender="Boy",Salary = 20000 },
        new Employees { FirstName = "Huynh", LastName = "Van",Gender="Ok",Salary = 4000 },
    };

            using (var package = new ExcelPackage(file))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(list, true);
                package.Save();
            }

            return ResultObject<string>.GetResult(0, "OK", downloadUrl);
        }

        [HttpGet("exportv2")]
        public async Task<IActionResult> ExportV2(CancellationToken cancellationToken)
        {
            // query data from database  
            await Task.Yield();
            var list = new List<Employees>()
    {
        new Employees { FirstName = "catcher",LastName="Anh" ,Gender="fd",Salary = 18 },
        new Employees {  FirstName = "catcherf",LastName="Anhd" ,Gender="fdf",Salary = 19 }
    };
            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(list, true);
                package.Save();
            }
            stream.Position = 0;
            string excelName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

            //return File(stream, "application/octet-stream", excelName);  
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }
    }
}

