using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Data;
using Task1.Models;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace Task1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BikeStoresController : ControllerBase
    {
        //create context variable
        private readonly BikeStoresContext _context;

        public BikeStoresController(BikeStoresContext context)
        {
            _context = context;
        }

        //to get all brands
        [HttpGet("GetAllBrand")]
        public async Task<IEnumerable<Brand>> GetBrands()
        {
            return await _context.Brands.ToListAsync();
        }

        //to get all products
        [HttpGet("GetAllProduct")]
        public async Task<IEnumerable<Product>> GetProducts()
        {
            return await _context.Products.ToListAsync();
        }

        //to get specific product
        [HttpGet("ProductById")]
        public async Task<IActionResult> GetProductById(int Id)
        {
            var product = await _context.Products.FindAsync(Id);
            return Ok(product);
        }

        //generate excel file from product table
        [HttpGet("ConvertToExcel")]
        public IActionResult GenerateExcelFile()
        {
            try
            {
                var data = _context.Products.ToList();

                //generate excel file using EPPlus
                byte[] excelFile = GenerateExcelFile(data);

                //return the excel file as a response
                return File(excelFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
                
            }

            catch (Exception ex)
            {
                return BadRequest("Error generating excel file: " + ex.Message);
            }
        }

        //upload excel then store in database
        [HttpPost("ExtractAndStore")]
        public IActionResult UploadExcelFile (IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("Please upload a valid excel file");
                }

                List<NewProduct> data = ExtractDataFromExcel(file);

                //save the extracted data to the new table in the database
                _context.NewProducts.AddRange(data);
                _context.SaveChanges();

                return Ok("Data successfully extracted and saved to the database");
            } 

            catch (Exception ex)
            {
                return BadRequest("Error processing the excel file: " + ex.Message);
            }
        }

        private byte[] GenerateExcelFile<T>(IEnumerable<T> data)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");            
                
                // Write column headers
                int columnIndex = 1;             
                foreach (var property in typeof(T).GetProperties())             
                {                 
                   worksheet.Cells[1, columnIndex].Value = property.Name;                 
                   columnIndex++;             
                }   
                    
                // Write data rows
                int rowIndex = 2;             
                foreach (var item in data)             
                {                 
                    columnIndex = 1;                 
                    foreach (var property in typeof(T).GetProperties())                 
                    {
                        worksheet.Cells[rowIndex, columnIndex].Value = property.GetValue(item);                    
                        columnIndex++;                 
                    }                 
                    rowIndex++;             
                }            
                    
                // Auto-fit columns for better visualization
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();             
                    
                // Save the Excel file to a memory stream
                return package.GetAsByteArray();         
            }   
        }

        private List<NewProduct> ExtractDataFromExcel(IFormFile file)
        {
            List<NewProduct> extractedData = new List<NewProduct>();

            using (var memoryStream = new MemoryStream())
            {
                file.CopyTo(memoryStream);
                using (ExcelPackage package = new ExcelPackage(memoryStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    int totalRows = worksheet.Dimension.Rows;

                    for (int row = 2; row <= totalRows; row++)
                    {
                        NewProduct newProduct = new NewProduct
                        {
                            ProductId = Int16.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            ProductName = worksheet.Cells[row, 2].Value.ToString(),
                            BrandId = Int16.Parse(worksheet.Cells[row, 3].Value.ToString()),
                            CategoryId = Int16.Parse(worksheet.Cells[row, 4].Value.ToString()),
                            ModelYear = Int16.Parse(worksheet.Cells[row, 5].Value.ToString()),
                            ListPrice = Convert.ToDecimal(worksheet.Cells[row, 6].Value.ToString()),
                            //Brand = worksheet.Cells[row, 7].Value?.ToString(),
                            //Category = worksheet.Cells[row, 8].Value?.ToString(),
                            //OrderItems = worksheet.Cells[row, 9].Value?.ToString(),
                            //Stocks = worksheet.Cells[row, 10].Value?.ToString(),
                        };

                        extractedData.Add(newProduct);
                    }
                }
            }

            return extractedData;
        }
    }
}
