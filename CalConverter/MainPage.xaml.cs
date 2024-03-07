using CalConverter.Lib;
using CommunityToolkit.Maui.Alerts;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Maui.Views;
using System.Text;
using System.Threading;

namespace CalConverter;

public partial class MainPage : ContentPage
{
    string file = string.Empty;


    private PickOptions options = new()
    {
        PickerTitle = "Please select preceptor schedule",
        FileTypes = new FilePickerFileType(
                new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.iOS, new[] { "com.microsoft.excel.xls" } }, // UTType values
                    { DevicePlatform.Android, new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } }, // MIME type
                    { DevicePlatform.WinUI, new[] { ".xlsx" } }, // file extension
                    { DevicePlatform.Tizen, new[] { "*/*" } },
                    { DevicePlatform.macOS, new[] { "com.microsoft.excel.xls" } }, // UTType values
                })
    };

    public MainPage()
    {
        InitializeComponent();
        ProcessButton.IsEnabled = false;
    }

    private async void OnOpenFileClicked(object sender, EventArgs e)
    {
        try
        {
            var result = await FilePicker.Default.PickAsync(options);
            if (result != null)
            {
                file = result.FullPath;
                FileNameLabel.Text = result.FileName;
                ProcessButton.IsEnabled = true;
            }

            return;
        }
        catch (Exception ex)
        {
            FileNameLabel.Text = "Invalid File Chosen!";
        }
    }

    private async void OnProcessClicked(object sender, EventArgs e)
    {
        var popup = new SpinnerPopup();
        popup.CanBeDismissedByTappingOutsideOfPopup = false;
        //await this.ShowPopupAsync(popup).ConfigureAwait(false);

        //await Task.Run(async () =>
        //{

            Parser parser = new Parser();
            Exporter exporter = new Exporter();
            try
            {
                var items = parser.ProcessFile(file, "Sheet1");
                foreach (var item in items)
                {
                    exporter.AddToCalendar(item);
                }
                using var stream = exporter.ToStream();
                var fileSaverResult = await FileSaver.Default.SaveAsync("preceptor.ics", stream);
                if (fileSaverResult.IsSuccessful)
                {
                    await Toast.Make($"The file was saved successfully to location: {fileSaverResult.FilePath}").Show();
                }
                else
                {
                    await Toast.Make($"The file was not saved successfully with error: {fileSaverResult.Exception.Message}").Show();
                }
            }
            catch (Exception ex)
            {
                Status.Text = ex.ToString();
            }
        //}).ConfigureAwait(false); ;
        //await popup.CloseAsync().ConfigureAwait(false); ;

    }
}

