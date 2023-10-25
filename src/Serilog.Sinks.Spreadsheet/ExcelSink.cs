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

    public async Task EmitBatchAsync(IEnumerable<LogEvent>? batch)
    {
        if (batch == null) return;
        var dic = new Dictionary<string, List<LogEvent>>();
        foreach(var logEvent in batch)
        {
            // Get log file name
            TextWriter buffer = new StringWriter(new StringBuilder());
            var fileName = _fileNameFactory(logEvent);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                SelfLog.WriteLine("Get a null log file name");
                continue;
            }

            new MessageTemplateTextFormatter(fileName).Format(logEvent, buffer);
            var name = buffer.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                SelfLog.WriteLine("No name specified for event");
                continue;
            }

            if (dic.TryGetValue(name, out var list))
            {
                list!.Add(logEvent);
                continue;
            }

            dic.Add(name, new List<LogEvent> { logEvent });
            
            var dir = Path.GetDirectoryName(name);
            if(string.IsNullOrWhiteSpace(dir)) continue;
            if (Directory.Exists(dir)) continue;
            
            try
            {
                Directory.CreateDirectory(dir);
            }
            catch (Exception exception)
            {
                SelfLog.WriteLine("Process events throw exception : {Message}\r\n{Exception}", exception.Message, exception);
            }
        }
        
        foreach(var item in dic)
        {
            await _handler.BatchAsync(item.Value.ToArray(), item.Key);
        }
    }

    public Task OnEmptyBatchAsync()
    {
        return Task.CompletedTask;
    }
}