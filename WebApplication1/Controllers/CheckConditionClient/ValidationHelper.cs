using NPOI.SS.UserModel;
using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;
using WebApplication1.Models;

namespace WebApplication1.Controllers.CheckConditionClient
{
    public class ValidationHelper
    {
        //Hàm kiểm tra điều kiện giá trị của từng row
        public class ValidationResult
        {
            public string Message { get; set; }
            public string PositionError { get; set; }
            public int Result { get; set; }
            public string ModifiedPhone { get; set; }
            public string Dob { get; set; }
        }

        public static ValidationResult CheckRowConditions(
            string contractNumber,
            List<Client> clientList,
            string phoneNumber,
            int row,
            string dob,
            string sheetName)
        {
            ValidationResult result = new ValidationResult();
            result.Message = "✔ Hợp đồng hợp lệ ! \n";

            //Kiểm tra mã hợp đồng
            if (contractNumber is "Null")
            {
                result.Message = "✗ Không có mã hợp đồng! \n";
                result.PositionError = $"Lỗi hợp đồng tại hàng {row + 1} của {sheetName} !\n";
                result.Result = 1;
            }
            else if (clientList.Any(c => c.C_IdContract == contractNumber))
            {
                result.Message = "✗ Chỉ lấy hợp đồng đầu tiên!\n";
                result.PositionError = $"Lỗi hàng {row + 1} của {sheetName} ! \n";
                result.Result = 1;
            }
            //Kiểm tra định dạng số điện thoại
            if (ValidateVietnameseMobileNumber(phoneNumber) == null)
            {
                result.Message += "⚠ Sai định dạng số điện thoại!\n";
                result.PositionError += $"Lỗi SĐT tại hàng {row + 1} của {sheetName} ! \n";
                result.ModifiedPhone = "Null";
            }
            else
            {
                // Nếu không có lỗi, đặt các giá trị mặc định cho các trường kết quả
                result.ModifiedPhone = phoneNumber;
            }
            //Kiểm tra định dạng kiểu ngày sinh trong file excel
            if (!string.IsNullOrEmpty(dob))
            {
                double b;
                //Kiểm tra parse value dob này có phải định dạng text ex: 37283, 
                if (double.TryParse(dob, out b))
                {
                    DateTime conv = DateTime.FromOADate(b);
                    //Kiểm tra xem ngày tháng năm có đúng định dạng hay không
                    if (conv >= DateTime.MinValue && conv <= DateTime.MaxValue)
                    {
                        result.Dob = conv.ToShortDateString();
                    }
                    else
                    {
                        result.Dob = "Null";
                        result.Message += "⚠ Không thể xác định ngày sinh!\n";
                    }
                }
                else
                {
                    // Nếu không thể parse thành số, kiểm tra xem có đúng định dạng "yyyy/MM/dd" hay không
                    if (DateTime.TryParseExact(dob, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime formattedDob))
                    {
                        result.Dob = formattedDob.ToShortDateString(); // trả về đúng định dạng 
                    }
                    else if (DateTime.TryParseExact(dob, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime formattedDob1)) // Kiểm tra trường có định dạng dd/MM/YYYY thì trả về đúng định dạng
                    {
                        result.Dob = formattedDob1.ToString("yyyy/MM/dd");
                    }
                    else if(DateTime.TryParse(dob, out DateTime parsedDate)) // Kiểm tra trường có định dạng kiểu datetime
                    {
                        result.Dob = parsedDate.ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        result.Dob = "Null"; // Hoặc "Null" khi không thể xác định kiểu dữ liệu date
                        result.Message += "⚠ Lỗi định dạng ngày sinh!\n";
                    }
                }
                //if (DateTime.TryParse(dob, out DateTime parsedDob))
                //{
                //    // Nếu thành công, kiểm tra liệu giá trị được parse có giống với ngày tháng không
                //    if (parsedDob.TimeOfDay == TimeSpan.Zero)
                //    {
                //        // Nếu giá trị được parse không chứa thông tin về thời gian (là ngày tháng), làm gì đó tương ứng
                //        dob = parsedDob.ToString("yyyy-MM-dd");
                //    }
                //    else
                //    {
                //        // Nếu giá trị được parse chứa thông tin về thời gian, đặt giá trị mặc định và thông báo lỗi
                //        result.Message += "⚠ Giá trị không phải là ngày tháng!\n";
                //        result.PositionError += $"Lỗi hàng {row} của {sheetName}!";
                //        result.Dob = "Null"; // Đặt giá trị mặc định hoặc để null tùy thuộc vào yêu cầu của bạn
                //    }
                //}
                //if (!DateTime.TryParseExact(dob, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime formattedDob))
                //{
                //    // Nếu không đúng định dạng, thực hiện chuyển đổi
                //    if (DateTime.TryParse(dob, out DateTime parsedDob1))
                //    {
                //        result.Dob = parsedDob1.ToString("yyyy-MM-dd");
                //    }
                //    else
                //    {
                //        // Nếu không thể chuyển đổi, đặt giá trị mặc định và thông báo lỗi
                //        result.Message += "⚠ Sai định dạng ngày sinh!\n";
                //        result.PositionError += $"Lỗi hàng {row} của {sheetName}!";
                //        result.Dob = "Null"; // Đặt giá trị mặc định hoặc để null tùy thuộc vào yêu cầu của bạn
                //    }
                //}
            }
            else
            {
                result.Dob = "Null";
            }
            return result;
        }
        //Hàm lấy giá trị của từng cột
        public static string GetCellValue(ISheet sheet, int row, int col)
        {


            if (col == -1) return "Null";

            var check = sheet.GetRow(row)?.GetCell(col)?.ToString()?.Trim();
            if (check is not null && check is not "")
            {
                return check.ToString();
            }
            return "Null";
        }
        //Hàm kiểm tra định dạng sdt VN
        public static string ValidateVietnameseMobileNumber(string number)
        {
            if (number is not null)
            {
                string pattern = @"^(09[0-9]|03[2-9]|07[0-9]|08[1-9]|05[6-9])[0-9]{7}$";

                return Regex.IsMatch(number, pattern) ? number : null;
            }
            return null;
        }
    }
}
