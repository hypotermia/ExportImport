using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;
using ExportImport.Models;
using Microsoft.EntityFrameworkCore;
using Rotativa.AspNetCore;

namespace ExportImport.Controllers
{
    public class ProductController : Controller
    {
        private readonly ExportImportContext _context;

        public ProductController(ExportImportContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var products = await _context.Products.ToListAsync();

            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Products");
                ws.Cells["A1"].LoadFromCollection(products, true);

                var stream = new MemoryStream();
                pck.SaveAs(stream);
                stream.Position = 0;

                string excelName = $"Products-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }
        [HttpPost]
        public async Task<IActionResult> ImportFromExcel(IFormFile file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (file != null && file.Length > 0)
            {
                using (var package = new ExcelPackage(file.OpenReadStream()))
                {
                    var ws = package.Workbook.Worksheets.First();
                    var rowCount = ws.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var product = new Product
                        {
                            Id = int.Parse(ws.Cells[row,1].Text),
                            Name = ws.Cells[row, 2].Text,
                            Price = decimal.Parse(ws.Cells[row, 3].Text)
                        };
                        _context.Products.Add(product);
                    }
                    await _context.SaveChangesAsync();
                }
            }
            return RedirectToAction(nameof(Index));
        }
        public async Task<IActionResult> Index()
        {
            var products = await _context.Products.ToListAsync();
            return View(products);
        }

        public async Task<IActionResult> ExportToPdf()
        {

            var products = await _context.Products.ToListAsync();
            return new ViewAsPdf("Index", products) { FileName = "Products.pdf" };
        }
    }
}
