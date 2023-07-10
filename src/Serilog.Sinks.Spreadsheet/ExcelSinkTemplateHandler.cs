using System.Collections.Concurrent;
using ClosedXML.Excel;
using Serilog.Debugging;
using Serilog.Events;

namespace Serilog.Sinks.Spreadsheet;

internal class ExcelSinkTemplateHandler : IExcelSinkHandler
{
    private readonly Func<LogEvent, string> _templateFactory;
    private string _logFileName;
    private ConcurrentDictionary<int, (string, XLDataType?, bool)>? _map;

    public ExcelSinkTemplateHandler(Func<LogEvent,string> templateFactory)
    {
        _logFileName = string.Empty;
        _templateFactory = templateFactory;
    }

    public Task BatchAsync(LogEvent[] events , string name)
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
                    var tpl = _templateFactory(events.First());
                    if (tpl == null)
                    {
                        SelfLog.WriteLine("Get template path is null:{@Event}",events.First());
                        return Task.CompletedTask;
                    }
                    workbook = new XLWorkbook(tpl);
                    worksheet = workbook.Worksheets.Worksheet(1);
                    _map ??= GetColumns(worksheet);
                    Append(worksheet, events, _map);
                    workbook.SaveAs(_logFileName);
                    return Task.CompletedTask;
                }
            }

            workbook = new XLWorkbook(_logFileName);
            worksheet = workbook.Worksheets.Worksheet(1);
            _map ??= GetColumns(worksheet);
            Append(worksheet, events, _map);
            workbook.Save();
        }
        finally
        {
            workbook?.Dispose();
        }
        
        return Task.CompletedTask;
    }

    private static ConcurrentDictionary<int, (string, XLDataType?, bool)> GetColumns(IXLWorksheet worksheet)
    {
        var fields = new ConcurrentDictionary<int, (string, XLDataType?, bool)>();
        var columns = worksheet.LastColumnUsed().ColumnNumber();
        for (var i = 1; i <= columns; i++)
        {
            var cell = worksheet.Cell(2, i);
            if (cell.DataType is XLDataType.Blank or XLDataType.Error) continue;

            var text = cell.GetText().Trim();
            var array = text.Split(':');

            if (array.Length == 1)
            {
                fields.TryAdd(i, (text, null, false));
                continue;
            }

            if ("FUN".Equals(array[0], StringComparison.OrdinalIgnoreCase))
            {
                fields.TryAdd(i, (array[1], null, true));
                continue;
            }

            if (Enum.IsDefined((typeof(XLDataType)), array[1]) && Enum.TryParse(array[1], true, out XLDataType type))
            {
                fields.TryAdd(i, (array[0], type, false));
                continue;
            }

            fields.TryAdd(i, (array[0], null, false));
            break;
        }

        return fields;
    }

    private static void Append(IXLWorksheet worksheet,LogEvent[] events, ConcurrentDictionary<int, (string, XLDataType?, bool)> map)
    {
        var rowNumber = worksheet.LastRowUsed().RowNumber();
        foreach(var logEvent in events)
        {
            rowNumber++;
            try
            {
                foreach(var tuple in map)
                {
                    var cell = worksheet.Cell(rowNumber , tuple.Key);
                    if (tuple.Value.Item3)
                    {
                        cell.SetFormulaA1(string.Format(tuple.Value.Item1, rowNumber));
                        continue;
                    }

                    if (!logEvent.Properties.TryGetValue(tuple.Value.Item1, out var val)) continue;

                    var value = val is StructureValue sv ? sv.ToString("l", null).Trim('"') : val.ToString().Trim('"');
                    if (tuple.Value.Item2.HasValue)
                    {
                        switch (tuple.Value.Item2.Value)
                        {
                            case XLDataType.Boolean:
                                if (bool.TryParse(value, out var b))
                                {
                                    cell.Value = b;
                                    continue;
                                }

                                break;
                            case XLDataType.Number:
                                if (double.TryParse(value, out var n))
                                {
                                    cell.Value = n;
                                    continue;
                                }

                                break;
                            case XLDataType.DateTime:
                                if (DateTime.TryParse(value, out var d))
                                {
                                    cell.Value = d;
                                    continue;
                                }

                                break;
                            case XLDataType.TimeSpan:
                                if (TimeSpan.TryParse(value, out var t))
                                {
                                    cell.Value = t;
                                    continue;
                                }

                                break;
                            case XLDataType.Text:
                            case XLDataType.Blank:
                            case XLDataType.Error:
                            default:
                                break;
                        }
                    }

                    cell.Value = value;
                }
            }
            catch (Exception exception)
            {
                SelfLog.WriteLine("Append row throw exception : {Message}\r\n{Exception}", exception.Message, exception);
            }
        }
    }
}
