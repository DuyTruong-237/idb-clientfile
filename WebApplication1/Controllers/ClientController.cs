using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
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
                }

                // Thực hiện các bước xử lý khác nếu cần

                return Ok("File uploaded successfully");
            }
            // Thêm các API khác ở đây nếu cần
        }
    }
}

