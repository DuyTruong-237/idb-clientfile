﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using WebApplication1.Data;
using WebApplication1.Models;
using System.Text.RegularExpressions;

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
        public async Task<IActionResult> UploadExcelFile(uploadFile file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (file == null || file.urlFile.Length == 0)
            {
                return BadRequest("File is null or empty");
            }

            //if (!Path.GetExtension(file.FileName).Equals(".xlsx", System.StringComparison.OrdinalIgnoreCase))
            //{
            //    return BadRequest("Invalid file format. Please upload a valid Excel file.");
            //}

            //Danh sách client
            List<Client> clientList = new List<Client>();

            string ValidateVietnameseMobileNumber(string number)
            {
                string pattern = @"^(09[0-9]|03[2-9]|07[0-9]|08[1-9]|05[6-9])[0-9]{7}$";

                return Regex.IsMatch(number, pattern) ? number : null;
            }
            string filePath = file.urlFile;
            if (filePath is null)
            {
                return NotFound("File not found");
            }

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                //await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    //Xác định tên của cột
                    Dictionary<string, int> columnMappings = new Dictionary<string, int>
                    {
                        { "Số hợp đồng", -1 },
                        { "Tên khách hàng", -1 },
                        { "CMND", -1 },
                        { "Ngày Sinh", -1 },
                        { "SDT", -1 },
                        { "Địa chỉ", -1 },
                        { "Tổng nợ", -1 },
                        { "Số tiền cần thanh toán hàng tháng", -1 }
                    };
                    //Xác định thứ tự các cột có trong file 
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[1, col].Value?.ToString().Trim();
                        if (columnMappings.ContainsKey(cellValue))
                        {
                            columnMappings[cellValue] = col;
                        }
                    }

                    // Kiểm tra tiêu đề các cột
                    if (columnMappings.Any(pair => pair.Value == -1))
                    {
                        Console.WriteLine("File không đáp ứng yêu cầu về tiêu đề cột.");
                        var respondData1 = new
                        {
                            agent = file.createBy,
                            dateImport = file.dateUpload,
                            messeage = "File không đáp ứng yêu cầu về tiêu đề cột!"
                        };

                        return Ok(respondData1);
                    }
                    int batchSize = 20000;
                    int totalRows = rowCount - 1; // Bỏ qua hàng đầu tiên vì nó là tiêu đề
                    int batches = (int)Math.Ceiling((double)totalRows / batchSize);
                    for (int batchIndex = 0;batchIndex < batches; batchIndex++)
                    {
                        int startRow = batchIndex * batchSize + 1; // Bắt đầu từ hàng thứ 2 (hàng đầu tiên là tiêu đề)
                        int endRow = Math.Min((batchIndex + 1) * batchSize, totalRows) + 1;

                        // Xử lý từ startRow đến endRow
                        for (int row = startRow+1; row <= endRow; row++)
                        {
                            // Thực hiện xử lý của bạn ở đây
                            var soHopDong = GetCellValue(worksheet, row, columnMappings["Số hợp đồng"]);
                            var name = GetCellValue(worksheet, row, columnMappings["Tên khách hàng"]);
                            var cmnd = GetCellValue(worksheet, row, columnMappings["CMND"]);
                            var dob = GetCellValue(worksheet, row, columnMappings["Ngày Sinh"]);
                            var phone = GetCellValue(worksheet, row, columnMappings["SDT"]);
                            var address = GetCellValue(worksheet, row, columnMappings["Địa chỉ"]);
                            var tongno = GetCellValue(worksheet, row, columnMappings["Tổng nợ"]);
                            var hangthang = GetCellValue(worksheet, row, columnMappings["Số tiền cần thanh toán hàng tháng"]);
                            var message = "✔ Hợp đồng hợp lệ";
                            var result = 0;
                            //Kiểm tra điều kiện
                            //Nếu hợp động null thì skip thực hiện qua row tiếp theo hoặc nếu trùng với số hợp đồng đã có từ trước sẽ lấy thằng đầu 
                            if (soHopDong == null)
                            {
                                message = "✗ Không có mã hợp đồng! \n ";
                                result = 1;
                                //continue;
                            }
                            else if (clientList.Any(c => c.C_IdContract == soHopDong))
                            {
                                message = "✗ Chỉ lấy hợp đồng đầu tiên! \n";
                                result = 1;
                                //continue;
                            }
                            if (ValidateVietnameseMobileNumber(phone) == null)
                            {
                                message += "⚠ Sai định dạng số điện thoại! \n";
                                phone = "";
                            }

                            var client = new Client
                            {
                                C_IdContract = soHopDong,
                                C_Name = name,
                                C_CMND = cmnd,
                                C_DayOfBirth = dob,
                                C_Phone = phone,
                                C_Address = address,
                                C_Totalliabilities = tongno,
                                C_AmountMonthly = hangthang,
                                messeage = message,
                                result = result
                            };
                            //Thêm vào danh sách 
                            clientList.Add(client);
                            //Add thông qua store procedure của SQL
                            try
                            {
                                if(client.result != 1)
                                {
                                    _context.Database.ExecuteSqlInterpolated($"CALL jwdb.CheckAndUpdateContract({client.Id},{client.C_IdContract}, {client.C_Name}, {client.C_CMND}, {client.C_DayOfBirth}, {client.C_Phone}, {client.C_Address}, {client.C_Totalliabilities}, {client.C_AmountMonthly})");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                    }

                    // Duyệt để lấy giá trị từng hàng 
                    //for (int row = 2; row <= rowCount; row++)
                    //{ 
                        
                    //    var soHopDong = GetCellValue(worksheet, row, columnMappings["Số hợp đồng"]);
                    //    var name = GetCellValue(worksheet, row, columnMappings["Tên khách hàng"]);
                    //    var cmnd = GetCellValue(worksheet, row, columnMappings["CMND"]);
                    //    var dob = GetCellValue(worksheet, row, columnMappings["Ngày Sinh"]);
                    //    var phone = GetCellValue(worksheet, row, columnMappings["SDT"]);
                    //    var address = GetCellValue(worksheet, row, columnMappings["Địa chỉ"]);
                    //    var tongno = GetCellValue(worksheet, row, columnMappings["Tổng nợ"]);
                    //    var hangthang = GetCellValue(worksheet, row, columnMappings["Số tiền cần thanh toán hàng tháng"]);
                    //    var message = "";
                    //    var result = 0;
                    //    //Kiểm tra điều kiện
                    //    //Nếu hợp động null thì skip thực hiện qua row tiếp theo hoặc nếu trùng với số hợp đồng đã có từ trước sẽ lấy thằng đầu 
                    //    if (soHopDong == null)
                    //    {
                    //        message += "Không có mã hợp đồng! \n ";
                    //        result = 1;
                    //        //continue;
                    //    }
                    //    else if (clientList.Any(c => c.C_IdContract == soHopDong))
                    //    {
                    //        message += "Chỉ lấy hợp đồng đầu tiên! \n";
                    //        result = 1;
                    //        //continue;
                    //    }
                    //    if (ValidateVietnameseMobileNumber(phone) == null)
                    //    {
                    //        message += "Sai định dạng số điện thoại! \n";
                    //        phone = "";
                    //    }

                    //    var client = new Client
                    //    {
                    //        C_IdContract = soHopDong,
                    //        C_Name = name,
                    //        C_CMND = cmnd,
                    //        C_DayOfBirth = dob,
                    //        C_Phone = phone,
                    //        C_Address = address,
                    //        C_Totalliabilities = tongno,
                    //        C_AmountMonthly = hangthang,
                    //        messeage = message,
                    //        result = result
                    //    };
                    //    //Thêm vào danh sách 
                    //    clientList.Add(client);
                    //    //Add thông qua store procedure của SQL
                    //    try
                    //    {
                    //        _context.Database.ExecuteSqlInterpolated($"CALL jwdb.CheckAndUpdateContract({client.Id},{client.C_IdContract}, {client.C_Name}, {client.C_CMND}, {client.C_DayOfBirth}, {client.C_Phone}, {client.C_Address}, {client.C_Totalliabilities}, {client.C_AmountMonthly})");
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        Console.WriteLine(ex.Message);
                    //    }

                    //}
                    //Dùng cơ chế batch để thực hiện cho 20k inser vào db
                    await _context.SaveChangesAsync();
                }
            }

            var respondData = new
            {
                Clients = clientList,
                agent = file.createBy,
                dateImport = file.dateUpload
            };

            return Ok(respondData);
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int col)
        {
            if (col == -1) return null;
            return worksheet.Cells[row, col].Value?.ToString().Trim();
        }

        // Thêm các API khác ở đây nếu cần
    }
}
