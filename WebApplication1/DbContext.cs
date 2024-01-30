namespace WebApplication1;

using Microsoft.EntityFrameworkCore;


public class AppDbContext : DbContext
{
    protected readonly IConfiguration Configuration;

    public AppDbContext(IConfiguration configuration)
    {
        Configuration = configuration;
    }

    protected override void OnConfiguring(DbContextOptionsBuilder options)
    {
        // connect to mysql with connection string from app settings
        var connectionString = Configuration.GetConnectionString("DefaultConnection");
        options.UseMySql(connectionString, ServerVersion.AutoDetect(connectionString));
    }

    public DbSet<app_fd_purchase_request> app_fd_purchase_request { get; set; }
    public DbSet<Client> app_fd_info_client { get; set; }
}