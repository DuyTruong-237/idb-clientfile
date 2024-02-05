using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Configuration;
using WebApplication1.Controllers.CheckConditionClient;
using WebApplication1.Models;

namespace WebApplication1.Controllers.Logic
{
    public class ExcelProcessor
    {
        private readonly IConfiguration _configuration;

        public ExcelProcessor(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public List<Client> ProcessExcelFile(string filePath, uploadFile file)
        {
            List<Client> clientList = new List<Client>();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = null ;

                // Kiểm tra định dạng của file Excel
                if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    // Định dạng xlsx
                    workbook = new XSSFWorkbook(stream);
                }
                else if (Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    // Định dạng xls
                    workbook = new HSSFWorkbook(stream);
                }
                //Lấy số lượng sheet đang có trong file
                if (workbook is not null)
                {
                    int sheetCount = workbook.NumberOfSheets;
                    for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++)
                    {
                        //ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                        ISheet sheet = workbook.GetSheetAt(sheetIndex);

                        // Lấy tên của sheet hiện tại
                        string sheetName = sheet.SheetName;
                        //Lấy số dòng của sheet hiện tại
                        int rowCount = sheet.PhysicalNumberOfRows;
                        //Lấy số cột của sheet hiện tại
                        int colCount = 0;
                        if (rowCount > 0)
                        {
                            colCount = sheet.GetRow(0)?.PhysicalNumberOfCells ?? 0;
                        }
                        //Xác định tên của cột bỏ vào các vị trí có trong từ điển từ appsetting.Json

                        var excelColumnMappings = _configuration.GetSection("ExcelColumnMappings").Get<Dictionary<string, string>>();
                        Dictionary<string, int> columnMappings = new Dictionary<string, int>();
                        //Import những trường từ trong file setting.Json vào columnMappings
                        foreach (var mapping in excelColumnMappings)
                        {
                            columnMappings.Add(mapping.Value, -1);
                        }
                        //{
                        //{ "Số hợp đồng", -1 },
                        //{ "Tên khách hàng", -1 },
                        //{ "CMND", -1 },
                        //{ "Ngày Sinh", -1 },
                        //{ "SDT", -1 },
                        //{ "Địa chỉ", -1 },
                        //{ "Tổng nợ", -1 },
                        //{ "Số tiền cần thanh toán hàng tháng", -1 }
                        //};
                        //Xác định thứ tự các cột có trong file 
                        for (int col = 0; col <= colCount; col++)
                        {
                            var cellValue = sheet.GetRow(0)?.GetCell(col)?.ToString().Trim();
                            if (!string.IsNullOrEmpty(cellValue) && columnMappings.ContainsKey(cellValue))
                            {
                                columnMappings[cellValue] = col;
                            }
                        }

                        // Kiểm tra xem sheet hiện tại có chứa các cột hợp lệ hay không
                        //var missingColumns = columnMappings.Where(pair => pair.Value == -1).Select(pair => pair.Key).ToList();
                        // Kiểm tra tiêu đề các cột
                        if (!columnMappings.ContainsKey("Số hợp đồng") || columnMappings["Số hợp đồng"] == -1)
                        {
                            Console.WriteLine("File không đáp ứng yêu cầu về tiêu đề cột.");
                            var client0 = new Client
                            {
                                C_IdContract = "NULL",
                                C_Name = "",
                                C_CMND = "",
                                C_DayOfBirth = "",
                                C_Phone = "",
                                C_Address = "",
                                C_Totalliabilities = "",
                                C_AmountMonthly = "",
                                result = 1,
                                messeage = $"✗ File {file.fileName} không đúng theo template!",
                                positionError = $"Lỗi thiếu cột Hợp đồng của {sheetName}"
                            };
                            clientList.Add(client0);
                            continue;
                            //return clientList;
                        }
                        int batchSize1 = 20000;
                        int totalRows = rowCount - 1; // Bỏ qua hàng đầu tiên vì nó là tiêu đề
                        int batches1 = (int)Math.Ceiling((double)totalRows / batchSize1);
                        for (int batchIndex = 0; batchIndex < batchSize1; batchIndex++)
                         
                        {
                            int startRow = batchIndex * batchSize1 + 1; // Bắt đầu từ hàng thứ 2 (hàng đầu tiên là tiêu đề)
                            int endRow = Math.Min((batchIndex + 1) * batchSize1, totalRows)+1;

                            // Xử lý từ startRow đến endRow
                            for (int row = startRow; row < endRow; row++)
                            {
                                var soHopDong = columnMappings["Số hợp đồng"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Số hợp đồng"]) : "Null";
                                var name = columnMappings["Tên khách hàng"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Tên khách hàng"]) : "";
                                var cmnd = columnMappings["CMND"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["CMND"]) : "";
                                var dob = columnMappings["Ngày Sinh"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Ngày Sinh"]) : "";
                                var phone = columnMappings["SDT"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["SDT"]) : "";
                                var address = columnMappings["Địa chỉ"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Địa chỉ"]) : "";
                                var tongno = columnMappings["Tổng nợ"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Tổng nợ"]) : "";
                                var hangthang = columnMappings["Số tiền cần thanh toán hàng tháng"] != -1 ? ValidationHelper.GetCellValue(sheet, row, columnMappings["Số tiền cần thanh toán hàng tháng"]) : "";
                                var message = "✔ Hợp đồng hợp lệ";
                                var positionError = "NULL";
                                var result = 0;
                                var validationResult = ValidationHelper.CheckRowConditions(soHopDong, clientList, phone, row, dob, sheetName);
                                var client = new Client
                                {
                                    C_IdContract = soHopDong,
                                    C_Name = name,
                                    C_CMND = cmnd,
                                    C_DayOfBirth = validationResult.Dob,
                                    C_Phone = validationResult.ModifiedPhone,
                                    C_Address = address,
                                    C_Totalliabilities = tongno,
                                    C_AmountMonthly = hangthang,
                                    messeage = validationResult.Message,
                                    result = validationResult.Result,
                                    positionError = validationResult.PositionError,
                                };
                                //Thêm vào danh sách 
                                clientList.Add(client);
                            }
                        }
                    }
                }
                
            }
            return clientList;
        }
    }
}
