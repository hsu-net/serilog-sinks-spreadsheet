using System.Text;
using Serilog.Debugging;
using Serilog.Events;
using Serilog.Formatting.Display;
using Serilog.Parsing;
using Serilog.Sinks.PeriodicBatching;

namespace Serilog.Sinks.Spreadsheet;

internal interface IExcelSinkHandler
{
    Task BatchAsync(LogEvent[] events,string name);
}

internal class ExcelSink : IBatchedLogEventSink
{
    private readonly Func<LogEvent, string> _fileNameFactory;
    private readonly IExcelSinkHandler _handler;
    
    public ExcelSink(Func<LogEvent,string> fileNameFactory,bool propertiesEnabled, string messageTemplate = ExcelSinkExtensions.DefaultTemplate)
    {
        _fileNameFactory = fileNameFactory ?? throw new ArgumentNullException(nameof(fileNameFactory));
        var tpl = new MessageTemplateParser().Parse(messageTemplate);
        _handler = new ExcelSinkHandler(tpl,propertiesEnabled);
    }

    public ExcelSink(Func<LogEvent,string> fileNameFactory,Func<LogEvent,string> templateFactory, string messageTemplate = ExcelSinkExtensions.DefaultTemplate)
    {
        _fileNameFactory = fileNameFactory ?? throw new ArgumentNullException(nameof(fileNameFactory));
        if (templateFactory == null) throw new ArgumentNullException(nameof(templateFactory));
        var tpl = new MessageTemplateParser().Parse(messageTemplate);
        _handler = new ExcelSinkTemplateHandler(tpl,templateFactory);
    }

    public Task EmitBatchAsync(IEnumerable<LogEvent>? batch)
    {
        var events = batch as LogEvent[] ?? batch?.ToArray();
        if (events == null || events.Length==0) return Task.CompletedTask;
        
        // Get log file name
        TextWriter buffer = new StringWriter(new StringBuilder());
        var fileName = _fileNameFactory(events.First());
        if (string.IsNullOrWhiteSpace(fileName))
        {
            SelfLog.WriteLine("Get a null log file name");
            return Task.CompletedTask;
        }
        
        new MessageTemplateTextFormatter(fileName).Format(events.First(), buffer);
        var name = buffer.ToString();
        if (string.IsNullOrWhiteSpace(name))
        {
            SelfLog.WriteLine("No name specified for event");
            return Task.CompletedTask;
        }
        
        var dir = Path.GetDirectoryName(name);
        if (!Directory.Exists(dir))
        {
            try
            {
                Directory.CreateDirectory(dir!);
            }
            catch (Exception exception)
            {
                SelfLog.WriteLine("Process events throw exception : {Message}\r\n{Exception}", exception.Message, exception);
                return Task.CompletedTask;
            }
        }
        
        return _handler.BatchAsync(events,name);
    }

    public Task OnEmptyBatchAsync()
    {
        return Task.CompletedTask;
    }
}