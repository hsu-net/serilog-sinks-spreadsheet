# Serilog.Sinks.Spreadsheet

[![dev](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/build.yml/badge.svg?branch=dev)](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/build.yml)
[![preview](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/deploy.yml/badge.svg?branch=preview)](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/deploy.yml)
[![main](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/deploy.yml/badge.svg?branch=main)](https://github.com/seayxu/serilog-sinks-spreadsheet/actions/workflows/deploy.yml)
[![Documentation](https://img.shields.io/badge/docs-wiki-blueviolet.svg)](https://github.com/serilog/serilog/wiki)
[![Nuke Build](https://img.shields.io/badge/nuke-build-yellow.svg)](https://github.com/nuke-build/nuke)

A [Serilog](https://github.com/serilog/serilog) sink that writes log events to any `Spreadsheet`.

## Package Version

| Name | Stable | Preview |
|---|---|---|
| Nuget | [![NuGet](https://img.shields.io/nuget/v/Serilog.Sinks.Spreadsheet?style=flat-square)](https://www.nuget.org/packages/Serilog.Sinks.Spreadsheet) | [![NuGet](https://img.shields.io/nuget/vpre/Serilog.Sinks.Spreadsheet?style=flat-square)](https://www.nuget.org/packages/Serilog.Sinks.Spreadsheet) |
| MyGet | [![MyGet](https://img.shields.io/myget/godsharp/v/Serilog.Sinks.Spreadsheet?style=flat-square&label=myget)](https://www.myget.org/feed/godsharp/package/nuget/Serilog.Sinks.Spreadsheet) | [![MyGet](https://img.shields.io/myget/godsharp/vpre/Serilog.Sinks.Spreadsheet?style=flat-square&label=myget)](https://www.myget.org/feed/godsharp/package/nuget/Serilog.Sinks.Spreadsheet) |

## Getting Started

  ```ps
  PM> Install-Package Serilog.Sinks.Spreadsheet
  ```

  ```csharp
  var tpl = Path.Combine(Environment.CurrentDirectory,"tpl.xlsx");
  Log.Logger = new LoggerConfiguration()
      .MinimumLevel.Verbose()
      .WriteTo.Excel()
      .WriteTo.Excel(tpl,"{Timestamp:yyyy-MM-dd}-tpl.xlsx")
      .CreateLogger();
  ```

## License

  MIT