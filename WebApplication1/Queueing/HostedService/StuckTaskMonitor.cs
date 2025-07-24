
using Microsoft.Data.SqlClient;

namespace WebApplicationAPI.Queueing.HostedService
{
    public class StuckTaskMonitor : BackgroundService
    {

        private readonly IServiceProvider _serviceProvider;
        private readonly IConfiguration _configuration;

        public StuckTaskMonitor(IServiceProvider serviceProvider,IConfiguration configuration)
        {
            _serviceProvider = serviceProvider;
            _configuration = configuration;
        }


        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while(!stoppingToken.IsCancellationRequested)
            {
                await Task.Delay(TimeSpan.FromMinutes(5),stoppingToken);

                using var scope = _serviceProvider.CreateScope();
                var connectionString = _configuration.GetConnectionString("DBConnection");

                using SqlConnection conn = new SqlConnection(connectionString);

                await conn.OpenAsync();

                string sql = @" UPDATE 
                              ";
            }

            throw new NotImplementedException();
        }
    }
}
