﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Threading.Tasks;
namespace WebApplication1.Controllers

{
    using global::WebApplication1.Data;
    using Microsoft.AspNetCore.Mvc;


    namespace WebApplication1.Controllers
    {
        [Route("api/[controller]")]
        [ApiController]
        public class PurchaseRequestsController : ControllerBase
        {
            private readonly AppDbContext _context;

            public PurchaseRequestsController(AppDbContext context)
            {
                _context = context;
            }

            // GET: api/PurchaseRequests
            [HttpGet]
            public async Task<ActionResult<IEnumerable<app_fd_purchase_request>>> GetPurchaseRequests()
            {
                return await _context.app_fd_purchase_request.ToListAsync();
            }

            // Thêm các API khác ở đây nếu cần
        }
    }
}

