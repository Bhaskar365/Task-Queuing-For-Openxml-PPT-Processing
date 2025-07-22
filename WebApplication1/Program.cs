using Microsoft.EntityFrameworkCore;
using WebApplication1.Context;
using WebApplication1.Repositories;
using WebApplicationAPI.Queueing;
using WebApplicationAPI.Queueing.HostedService;
using WebApplicationAPI.Seeding;
using WebApplicationAPI.TaskTracking;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddDbContext<AppDBContext>(option => option.UseSqlServer(builder.Configuration.GetConnectionString("DBConnection")));

builder.Services.AddScoped<IDataRepository, DataRepository>();

builder.Services.AddCors(option =>
{
    option.AddPolicy("AllowAll",p => p.AllowAnyOrigin().AllowAnyHeader().AllowCredentials());
});

builder.Services.AddSingleton<IBackgroundTaskQueue,BackgroundTaskQueue>();

builder.Services.AddSingleton<ITaskStatusTracker, TaskStatusTracker>();

builder.Services.AddHostedService<QueuedHostedService>();

var app = builder.Build();

using (var scope = app.Services.CreateScope())
{
    var dbContext = scope.ServiceProvider.GetRequiredService<AppDBContext>();
    var seeder = new CustomSeeder(dbContext);
    await seeder.SeedAsync();
}

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
