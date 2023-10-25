// See https://aka.ms/new-console-template for more information
using Serilog;
Console.WriteLine("Hello, Serilog.Sinks.Spreadsheet!"); 
var tpl = Path.Combine(Environment.CurrentDirectory,"tpl.xlsx");

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Verbose()
    .WriteTo.Console()
    //.WriteTo.Excel()
    .WriteTo.Excel(tpl,"{Timestamp:yyyy-MM-dd}-tpl.xlsx")
    .CreateLogger();

for (var i = 0; i < 10; i++)
{
    Log.Information("{Id},{Name:l},{Age}", i + 1, "N" + i, i + 2);
}

Log.CloseAndFlush();