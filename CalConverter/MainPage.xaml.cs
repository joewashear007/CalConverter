using CalConverter.Lib;
using CommunityToolkit.Maui.Alerts;
using CommunityToolkit.Maui.Storage;


namespace CalConverter;

public partial class MainPage : ContentPage
{
    private readonly Parser parser;
    private readonly Exporter exporter;



    public MainPage(MainPageViewModel viewModel)
    {
        InitializeComponent();
        BindingContext = viewModel;       
    }

   

    
}

