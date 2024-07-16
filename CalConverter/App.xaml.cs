using MetroLog.Maui;

namespace CalConverter;

public partial class App : Application
{
    public App()
    {
        InitializeComponent();

        MainPage = new AppShell();


        LogController.InitializeNavigation(
            page => MainPage!.Navigation.PushModalAsync(page),
            () => MainPage!.Navigation.PopModalAsync());
    }
}
