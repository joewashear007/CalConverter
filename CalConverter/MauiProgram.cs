using CalConverter.Lib;
using CalConverter.Lib.Parsers;
using CommunityToolkit.Maui;
using MetroLog.MicrosoftExtensions;
using Microsoft.Extensions.Logging;

namespace CalConverter;
public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();

        builder.Logging.AddConsole();
        builder.Services.AddTransient<BaseParser, PreceptorSchedule>();
        builder.Services.AddTransient<BaseParser, DidacticSchedule>();
        builder.Services.AddTransient<PreceptorSchedule>();
        builder.Services.AddTransient<DidacticSchedule>();
        builder.Services.AddTransient<Exporter>();

        builder
            .UseMauiApp<App>()
            .UseMauiCommunityToolkit()
            .ConfigureFonts(fonts =>
            {

                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

        builder.Services.AddSingleton<MainPageViewModel>();
        builder.Services.AddSingleton<MainPage>();
        builder.Services.AddTransient(a => FilePicker.Default);
        builder.Logging
#if DEBUG
                   .AddTraceLogger(
                       options =>
                       {
                           options.MinLevel = LogLevel.Trace;
                           options.MaxLevel = LogLevel.Critical;
                       }) // Will write to the Debug Output
#endif
                   .AddInMemoryLogger(
                       options =>
                       {
                           options.MaxLines = 1024;
                           options.MinLevel = LogLevel.Debug;
                           options.MaxLevel = LogLevel.Critical;
                       })
#if RELEASE
            .AddStreamingFileLogger(
                options =>
                {
                    options.RetainDays = 2;
                    options.FolderPath = Path.Combine(
                        FileSystem.CacheDirectory,
                        "MetroLogs");
                })
#endif
                   .AddConsoleLogger(
                       options =>
                       {
                           options.MinLevel = LogLevel.Information;
                           options.MaxLevel = LogLevel.Critical;
                       }); 

        return builder.Build();
    }
}
