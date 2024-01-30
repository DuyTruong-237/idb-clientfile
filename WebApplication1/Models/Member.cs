using Microsoft.EntityFrameworkCore;

namespace WebApplication1.Models
{

    public class Member
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    // DbContext


}
