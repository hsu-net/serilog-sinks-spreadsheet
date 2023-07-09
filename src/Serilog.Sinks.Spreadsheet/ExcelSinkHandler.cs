using System.Collections.Concurrent;
using System.Text;
using ClosedXML.Excel;
using Serilog.Debugging;
using Serilog.Events;
using Serilog.Formatting.Display;
using Serilog.Parsing;

namespace Serilog.Sinks.Spreadsheet;

internal class ExcelSinkHandler : IExcelSinkHandler
{
    private readonly bool _propertiesEnabled;
    private string _logFileName;
    private readonly MessageTemplate _messageTemplate;
    private readonly MessageTemplateTextFormatter? _timestampTemplate;
    private readonly MessageTemplateTextFormatter? _levelTemplate;
    private readonly MessageTemplateTextFormatter? _exceptionTemplate;
    private ConcurrentDictionary<int, PropertyToken>? _map;

    public ExcelSinkHandler(MessageTemplate messageTemplate,bool propertiesEnabled)
    {
        _logFileName = string.Empty;
        _propertiesEnabled = propertiesEnabled;
        _messageTemplate = messageTemplate;
        foreach(var token in _messageTemplate.Tokens)
        {
            if (token is not PropertyToken property) continue;
            switch (property.PropertyName)
            {
                case OutputProperties.TimestampPropertyName:
                    _timestampTemplate = new MessageTemplateTextFormatter(property.ToString());
                    break;
                case OutputProperties.LevelPropertyName:
                    _levelTemplate = new MessageTemplateTextFormatter(property.ToString());
                    break;
                case OutputProperties.ExceptionPropertyName:
                    _exceptionTemplate = new MessageTemplateTextFormatter(property.ToString());
                    break;
            }
        }
    }
    
    public Task BatchAsync(LogEvent[] events, string name)
    {
        XLWorkbook? workbook=null;
        try
        {
            IXLWorksheet worksheet;
            if (name != _logFileName)
            {
                _logFileName = name;
                if (!Utils.Exists(_logFileName))
                {
                    workbook = new XLWorkbook();
                    worksheet = workbook.Worksheets.Add(1);
                    _map ??= GetColumns(events.First(), _messageTemplate,_propertiesEnabled);
                    Append(worksheet, _map);
                    Append(worksheet, events, _map);
                    workbook.SaveAs(_logFileName);
                    return Task.CompletedTask;
                }
            }
            
            workbook = new XLWorkbook(_logFileName);
            worksheet = workbook.Worksheets.Worksheet(1);
            _map ??= GetColumns(events.First(),_messageTemplate,_propertiesEnabled);
            Append(worksheet, events, _map);
            workbook.Save();
        }
        finally
        {
            workbook?.Dispose();
        }
        
        return Task.CompletedTask;
    }

    private static ConcurrentDictionary<int, PropertyToken> GetColumns(LogEvent logEvent,MessageTemplate messageTemplate,bool propertiesEnabled)
    {
        var fields = new ConcurrentDictionary<int, PropertyToken>();
        var columnIndex = 1;

        foreach(var token in messageTemplate.Tokens)
        {
            if (token is not PropertyToken property || property.PropertyName == OutputProperties.NewLinePropertyName) continue;
            if (property.PropertyName == OutputProperties.MessagePropertyName && propertiesEnabled) TryAdd(logEvent.MessageTemplate);
            fields.TryAdd(columnIndex++, property);
        }

        return fields;

        void TryAdd(MessageTemplate template)
        {
            foreach(var token in template.Tokens)
            {
                if (token is not PropertyToken property) continue;
                fields.TryAdd(columnIndex++, property);
            }
        }
    }

    private void Append(IXLWorksheet worksheet,LogEvent[] events, ConcurrentDictionary<int, PropertyToken> map)
    {
        var rowNumber = worksheet.LastRowUsed().RowNumber();
        foreach(var logEvent in events)
        {
            rowNumber++;
            try
            {
                foreach(var mp in map)
                {
                    using var writer = new StringWriter(new StringBuilder());
                    switch (mp.Value.PropertyName)
                    {
                        case OutputProperties.TimestampPropertyName:
                            _timestampTemplate?.Format(logEvent,writer);
                            break;
                        case OutputProperties.LevelPropertyName:
                            _levelTemplate?.Format(logEvent,writer);
                            break;
                        case OutputProperties.ExceptionPropertyName:
                            _exceptionTemplate?.Format(logEvent,writer);
                            break;
                        case OutputProperties.MessagePropertyName:
                            logEvent.RenderMessage(writer);
                            break;
                        default:
                            mp.Value.Render(logEvent.Properties, writer);
                            break;
                    }

                    worksheet.Cell(rowNumber, mp.Key).Value = writer.ToString().Trim('"');
                }
            }
            catch (Exception exception)
            {
                SelfLog.WriteLine("Append row throw exception : {Message}\r\n{Exception}", exception.Message, exception);
            }
        }
    }
    
    private static void Append(IXLWorksheet worksheet, ConcurrentDictionary<int, PropertyToken> map)
    {
        foreach(var mp in map)
        {
            worksheet.Cell(1, mp.Key).Value = mp.Value.PropertyName;
        }
    }
}
