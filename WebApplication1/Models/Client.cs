namespace WebApplication1.Models
{
    public class Client
    {
        public string Id { get; set; }
        public DateTime DateCreated { get; set; }
        public DateTime DateModified { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedByName { get; set; }
        public string ModifiedBy { get; set; }
        public string ModifiedByName { get; set; }
        public string C_Totalliabilities { get; set; }
        public string C_Address { get; set; }
        public string C_AmountMonthly { get; set; }
        public string C_Phone { get; set; }
        public string C_IdContract { get; set; }
        public string C_Name { get; set; }
        public string C_DayOfBirth { get; set; }
        public string C_CMND { get; set; }
        public ICollection<Member> members { get; set; }
    }
}
