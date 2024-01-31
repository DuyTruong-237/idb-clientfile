using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Threading.Tasks;
namespace WebApplication1.Controllers

{
    using global::WebApplication1.Data;
    using global::WebApplication1.Models;
    using Microsoft.AspNetCore.Mvc;


    namespace WebApplication1.Controllers
    {
        [Route("api/[controller]")]
        [ApiController]
        public class ClientsController : ControllerBase
        {
            private readonly AppDbContext _context;

            public ClientsController(AppDbContext context)
            {
                _context = context;
            }

            // GET: api/PurchaseRequests
            [HttpGet]
            public async Task<ActionResult<IEnumerable<Client>>> GetPurchaseRequests()
            {
                return await _context.app_fd_info_client.ToListAsync();
            }
            [HttpPost("upload")]
            public async Task<IActionResult> UploadExcelFile(IFormFile file)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (file == null || file.Length == 0)
                {
                    return BadRequest("File is null or empty");
                }

                if (!Path.GetExtension(file.FileName).Equals(".xlsx", System.StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest("Invalid file format. Please upload a valid Excel file.");
                }

                // Đọc dữ liệu từ file Excel
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    // Bạn có thể thực hiện xử lý dữ liệu ở đây, ví dụ: đọc dữ liệu từ Excel sử dụng thư viện như EPPlus.
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var soHopDong = worksheet.Cells[row, 1].Value?.ToString().Trim();
                            var name = worksheet.Cells[row, 2].Value?.ToString().Trim();
                            var cmnd = worksheet.Cells[row, 3].Value?.ToString().Trim();
                            var dob = worksheet.Cells[row, 4].Value?.ToString().Trim();
                            var phone = worksheet.Cells[row, 5].Value?.ToString().Trim();
                            var address = worksheet.Cells[row, 6].Value?.ToString().Trim();
                            var tongno = worksheet.Cells[row, 7].Value?.ToString().Trim();
                            var hangthang = worksheet.Cells[row, 8].Value?.ToString().Trim();
                            // Tạo đối tượng từ dữ liệu đọc được
                            var yourDataObject = new Client
                            {
                                C_IdContract = soHopDong,
                                C_Name = name,
                                C_CMND=cmnd,
                                C_DayOfBirth = dob,
                                C_Phone = phone,
                                C_Address=address,
                                C_Totalliabilities=tongno,
                                C_AmountMonthly=hangthang
                               
                            };

                            // Lưu vào cơ sở dữ liệu
                            _context.app_fd_info_client.Add(yourDataObject);
                        }

                        await _context.SaveChangesAsync();
                    }
                }

                // Thực hiện các bước xử lý khác nếu cần

                return Ok("File uploaded successfully");
            }
            // Thêm các API khác ở đây nếu cần
        }
    }
}

