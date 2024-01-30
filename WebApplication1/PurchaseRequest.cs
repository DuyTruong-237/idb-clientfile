using System;
namespace WebApplication1
{
    public class app_fd_purchase_request
    {
        public string Id { get; set; }
        public DateTime DateCreated { get; set; }
        public DateTime DateModified { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedByName { get; set; }
        public string ModifiedBy { get; set; }
        public string ModifiedByName { get; set; }
        public string C_Name { get; set; }
       
        public string C_Items { get; set; }
        public string C_Remarks { get; set; }
        public string C_Category { get; set; }
    }
}
