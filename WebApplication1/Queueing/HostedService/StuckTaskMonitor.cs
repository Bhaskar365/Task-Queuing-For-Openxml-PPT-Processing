
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
                await Task.Delay(TimeSpan.FromMinutes(3),stoppingToken);

                using var scope = _serviceProvider.CreateScope();
                var connectionString = _configuration.GetConnectionString("DBConnection");

                using SqlConnection conn = new SqlConnection(connectionString);

                await conn.OpenAsync();

                /*--- 5 minute status logic --*/
                string sql = @"UPDATE 
                                TaskLoggingTable SET 
                                CurrentStatus = 'Fail',
                                CompletedOn= GETUTCDATE()
                             WHERE
                                CurrentStatus = 'Processing' OR CurrentStatus = 'Fail'
                                AND CompletedOn IS NULL
                                AND DATEDIFF(MINUTE,CreatedOn,GETUTCDATE())>=5              
                            ";

                using SqlCommand cmd = new SqlCommand(sql, conn);
                await cmd.ExecuteNonQueryAsync();
            }

            throw new NotImplementedException();
        }
    }
}
