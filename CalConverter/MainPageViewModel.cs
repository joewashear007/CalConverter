using CalConverter.Lib;
using CommunityToolkit.Maui.Alerts;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace CalConverter;
public class MainPageViewModel : INotifyPropertyChanged
{
    private bool isProcessing;
    private DateTime endDate;
    private DateTime startDate;
    private FileResult? file;
    private readonly Parser parser;
    private readonly Exporter exporter;
    private readonly IFilePicker filePicker;
    private string statusMessages = "";

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
    private bool filePerPerson;

    public MainPageViewModel(Parser parser, Exporter exporter, IFilePicker filePicker)
    {

        File = null;
        this.parser = parser;
        this.exporter = exporter;
        this.filePicker = filePicker;
        this.StartDate = DateTime.Today;
        this.EndDate = DateTime.Today.AddMonths(6);
        LoadFile = new AsyncRelayCommand(
            execute: async () =>
            {
                IsProcessing = true;
                try
                {

                    var result = await filePicker.PickAsync(options);
                    if (result != null)
                    {
                        File = result;
                    }
                }
                catch (Exception ex)
                {
                    StatusMessages = "Invalid File Chosen!";
                }
                finally
                {
                    IsProcessing = false;
                    ProcessFile.NotifyCanExecuteChanged();
                    LoadFile.NotifyCanExecuteChanged();
                }
            },
            canExecute: () =>
            {
                return !IsProcessing;
            });

        ProcessFile = new AsyncRelayCommand(
            execute: async () =>
            {
                IsProcessing = true;
                try
                {
                    exporter.Options.ExportStartDate = DateOnly.FromDateTime(StartDate.Date);
                    exporter.Options.ExportEndDate = DateOnly.FromDateTime(EndDate.Date);
                    exporter.Options.FilePerPerson = filePerPerson;

                    var items = parser.ProcessFile(File.FullPath, "Sheet1");
                    foreach (var item in items)
                    {
                        exporter.AddToCalendar(item);
                    }
                    if (exporter.Options.FilePerPerson)
                    {

                        var folderPickerResult = await FolderPicker.Default.PickAsync();
                        if (folderPickerResult.IsSuccessful)
                        {
                            foreach (var key in exporter.Calendars.Keys.Where(q => q != "ALL"))
                            {
                                using var stream = exporter.ToStream(key);
                                string filename = Path.Join(folderPickerResult.Folder.Path, $"preceptor-{key}.ics");
                                using var sw = new FileStream(filename, FileMode.OpenOrCreate);
                                stream.Seek(0, SeekOrigin.Begin);
                                await stream.CopyToAsync(sw);
                            }
                            await Toast.Make($"The file was saved successfully to location: {folderPickerResult.Folder.Path}").Show();
                        }
                        else
                        {
                            await Toast.Make($"The file was not saved successfully with error: {folderPickerResult.Exception.Message}").Show();
                        }
                    }
                    else
                    {
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
                }
                catch (Exception ex)
                {
                    StatusMessages = "Error!" ;
                    Exception? e = ex;
                    do
                    {
                        StatusMessages += Environment.NewLine + e.Message;
                        e = e.InnerException;
                    } while (e != null);

                    StatusMessages += Environment.NewLine + Environment.NewLine + ex.StackTrace;
                    File = null;
                }
                finally
                {
                    ProcessFile?.NotifyCanExecuteChanged();
                    LoadFile.NotifyCanExecuteChanged();
                    IsProcessing = false;
                }
            },
            canExecute: () =>
            {
                return !IsProcessing && File != null;
            });
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public bool IsProcessing
    {
        set { SetProperty(ref isProcessing, value); }
        get { return isProcessing; }
    }

    public FileResult? File
    {
        set
        {
            SetProperty(ref file, value);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileName)));
        }
        get { return file; }
    }

    public string FileName
    {
        get { return file?.FileName ?? "No File Loaded"; }
    }

    public DateTime StartDate
    {
        set { SetProperty(ref startDate, value); }
        get { return startDate; }
    }
    public DateTime EndDate
    {
        set { SetProperty(ref endDate, value); }
        get { return endDate; }
    }

    public string StatusMessages
    {
        set { SetProperty(ref statusMessages, value); }
        get { return statusMessages; }
    }

    public bool FilePerPerson
    {
        set { SetProperty(ref filePerPerson, value); }
        get { return filePerPerson; }
    }

    public bool ValidFileLoaded { get; set; } = false;


    public AsyncRelayCommand LoadFile { get; set; }
    public AsyncRelayCommand ProcessFile { get; set; }


    bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
    {
        if (Equals(storage, value))
            return false;

        storage = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }


}
