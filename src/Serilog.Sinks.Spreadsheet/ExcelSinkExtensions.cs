using Serilog.Configuration;
using Serilog.Events;
using Serilog.Sinks.PeriodicBatching;
using Serilog.Sinks.Spreadsheet;
// ReSharper disable CheckNamespace
// ReSharper disable MemberCanBePrivate.Global

namespace Serilog;

/// <summary>
/// Extension methods for <see cref="ExcelSink"/>.
/// </summary>
public static class ExcelSinkExtensions
{
    internal const string DefaultTemplate = "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
    internal const string DefaultLogName = "{Timestamp:yyyy-MM-dd}.xlsx";

    /// <summary>
    /// Write log event to excel without the specified template
    /// </summary>
    /// <param name="loggerSinkConfiguration"></param>
    /// <param name="fileName">The file name can be use log event properties.</param>
    /// <param name="propertiesEnabled">If write each properties to excel.</param>
    /// <param name="messageTemplate">Custom output columns message template,default is <see cref="DefaultTemplate"/>.</param>
    /// <param name="pbOptions"></param>
    /// <param name="minimumLevel"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static LoggerConfiguration Excel(this LoggerSinkConfiguration loggerSinkConfiguration,
        string fileName = DefaultLogName,
        bool propertiesEnabled = true,
        string messageTemplate = DefaultTemplate,
        PeriodicBatchingSinkOptions? pbOptions = null,
        LogEventLevel minimumLevel = LogEventLevel.Verbose
    )
    {
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
        return loggerSinkConfiguration.Excel(_ => fileName, propertiesEnabled, messageTemplate, pbOptions, minimumLevel);
    }

    /// <summary>
    /// Write log event to excel with the specified template
    /// </summary>
    /// <param name="loggerSinkConfiguration"></param>
    /// <param name="template"></param>
    /// <param name="fileName"></param>
    /// <param name="options"></param>
    /// <param name="minimumLevel"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static LoggerConfiguration Excel(this LoggerSinkConfiguration loggerSinkConfiguration,
        string template,
        string fileName = DefaultLogName,
        PeriodicBatchingSinkOptions? options = null,
        LogEventLevel minimumLevel = LogEventLevel.Verbose
    )
    {
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
        return loggerSinkConfiguration.Excel(template, _ => fileName, options, minimumLevel);
    }

    /// <summary>
    /// Write log event to excel without the specified template
    /// </summary>
    /// <param name="loggerSinkConfiguration"></param>
    /// <param name="fileNameFactory"></param>
    /// <param name="propertiesEnabled"></param>
    /// <param name="messageTemplate"></param>
    /// <param name="options"></param>
    /// <param name="minimumLevel"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static LoggerConfiguration Excel(this LoggerSinkConfiguration loggerSinkConfiguration,
        Func<LogEvent, string> fileNameFactory,
        bool propertiesEnabled = true,
        string messageTemplate = DefaultTemplate,
        PeriodicBatchingSinkOptions? options = null,
        LogEventLevel minimumLevel = LogEventLevel.Verbose
    )
    {
        if (fileNameFactory == null) throw new ArgumentNullException(nameof(fileNameFactory));
        return loggerSinkConfiguration.Sink(
            new PeriodicBatchingSink(new ExcelSink(fileNameFactory, propertiesEnabled, messageTemplate), options ?? DefaultPeriodicBatchingOptions)
            , minimumLevel);
    }

    /// <summary>
    /// Write log event to excel with the specified template
    /// </summary>
    /// <param name="loggerSinkConfiguration"></param>
    /// <param name="fileNameFactory"></param>
    /// <param name="template"></param>
    /// <param name="minimumLevel"></param>
    /// <param name="options"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static LoggerConfiguration Excel(this LoggerSinkConfiguration loggerSinkConfiguration,
        string template,
        Func<LogEvent, string> fileNameFactory,
        PeriodicBatchingSinkOptions? options = null,
        LogEventLevel minimumLevel = LogEventLevel.Verbose
    )
    {
        if (template == null) throw new ArgumentNullException(nameof(template));
        if (fileNameFactory == null) throw new ArgumentNullException(nameof(fileNameFactory));
        return loggerSinkConfiguration.Sink(
            new PeriodicBatchingSink(new ExcelSink(fileNameFactory, template), options ?? DefaultPeriodicBatchingOptions)
            , minimumLevel);
    }

    static readonly PeriodicBatchingSinkOptions DefaultPeriodicBatchingOptions = new()
    {
        BatchSizeLimit = 50,
        Period = TimeSpan.FromSeconds(5)
    };
}
