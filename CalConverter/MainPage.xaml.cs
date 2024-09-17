using CalConverter.Lib;
using CalConverter.Lib.Parsers;
using CommunityToolkit.Maui.Alerts;
using CommunityToolkit.Maui.Storage;


namespace CalConverter;

public partial class MainPage : ContentPage
{
    public MainPage(MainPageViewModel viewModel)
    {
        InitializeComponent();
        BindingContext = viewModel;       
    }    
}

