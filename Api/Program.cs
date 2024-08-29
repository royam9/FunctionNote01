using Services;
using Services.Interfaces;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle

# region Swagger
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
# endregion

# region Unit of Service
builder.Services.AddScoped<IConvertHtmlToPdfService, ConvertHtmlToPdfService>();
builder.Services.AddScoped<IReplaceHtmlTextService, ReplaceHtmlTextService>();
builder.Services.AddScoped<IReplaceWordTextService, ReplaceWordTextService>();
# endregion

var app = builder.Build();

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
