using CalConverter.Lib;
using CommunityToolkit.Maui;
using Microsoft.Extensions.Logging;

namespace CalConverter;
public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();

        builder.Logging.AddConsole();
        builder.Services.AddSingleton<Parser>();
        builder.Services.AddSingleton<Exporter>();

        builder
            .UseMauiApp<App>()
            .UseMauiCommunityToolkit()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

#if DEBUG
		builder.Logging.AddDebug();
#endif

        return builder.Build();
    }
}
