using CalConverter.Lib;
using CalConverter.Lib.Parsers;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.Input;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace CalConverter;
public partial class MainPageViewModel : INotifyPropertyChanged
{
    private bool isProcessing;
    private DateTime endDate;
    private DateTime startDate;
    private FileResult? file;
    private IServiceProvider serviceProvider;
    private IFilePicker filePicker;
    private string statusMessages = "";
    private string parserName = "DidacticSchedule";

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
    private bool filePerPerson = true;
    private bool exportAdminTime = false;
    private List<string> sheetNames = [];
    private string sheet = "Sheet1";

    public MainPageViewModel(IServiceProvider serviceProvider, IFilePicker filePicker)
    {

        File = null;
        this.serviceProvider = serviceProvider;
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
                        SheetNames = Utils.GetSheetNames(result.FullPath);
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
                StatusMessages = $"Staring to Process File '{FileName}' ...";
                IsProcessing = true;
                try
                {

                    Type parserType = Type.GetType("CalConverter.Lib.Parsers." + Parser + ", CalConverter.Lib");
                    if(parserType == null)
                    {
                        throw new InvalidOperationException($"Failed to type of parser!'");
                    }



                    var exporter = serviceProvider.GetRequiredService<Exporter>();
                    BaseParser parser = serviceProvider.GetRequiredService(parserType) as BaseParser;
                    exporter.Options.ExportStartDate = DateOnly.FromDateTime(StartDate.Date);
                    exporter.Options.ExportEndDate = DateOnly.FromDateTime(EndDate.Date);
                    exporter.Options.FilePerPerson = filePerPerson;
                    exporter.Options.ExportAdminTime = exportAdminTime;

                    var items = parser.ProcessFile(File.FullPath, Sheet);
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
                                string cleanFileName = CleanStringRegex().Replace(key, "");
                                string filename = Path.Join(folderPickerResult.Folder.Path, $"preceptor-{cleanFileName}.ics");
                                using var sw = new FileStream(filename, FileMode.OpenOrCreate);
                                stream.Seek(0, SeekOrigin.Begin);
                                await stream.CopyToAsync(sw);
                                StatusMessages += $"{Environment.NewLine}Saving Generated File: '{filename}'";
                            }
                        }
                        else
                        {
                            StatusMessages += $"{Environment.NewLine}Failed to get Folder to Save too!'";
                        }
                    }
                    else
                    {
                        using var stream = exporter.ToStream();
                        var fileSaverResult = await FileSaver.Default.SaveAsync("preceptor.ics", stream);
                        StatusMessages += $"{Environment.NewLine}Saving Generated File: '{fileSaverResult.FilePath}'";
                        if (!fileSaverResult.IsSuccessful)
                        {
                            StatusMessages += $"{Environment.NewLine}The file was not saved successfully with error: {fileSaverResult.Exception.Message}";
                        }
                        
                    }
                    StatusMessages += $"{Environment.NewLine}Process Complete!";
                }
                catch (Exception ex)
                {
                    StatusMessages += $"{Environment.NewLine}================ Error! ==========================";
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
                    IsProcessing = false;
                    LoadFile.NotifyCanExecuteChanged();
                    ProcessFile?.NotifyCanExecuteChanged();
                }
            },
            canExecute: () =>
            {
                return !IsProcessing && FileLoaded && !string.IsNullOrEmpty(Sheet);
            });
    }

    public event PropertyChangedEventHandler PropertyChanged;
   
    public List<string> ParserNames
    {
        get {
            return Utils.GetParserNames();
        }
    }

    public string Parser
    {
        set { SetProperty(ref parserName, value); }
        get { return parserName; }
    }

    public bool IsProcessing
    {
        set { 
            SetProperty(ref isProcessing, value);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileLoaded)));
        }
        get { return isProcessing; }
    }

    public List<string> SheetNames
    {
        set { SetProperty(ref sheetNames, value); }
        get { return sheetNames; }
    }

    public string Sheet
    {
        set { 
            SetProperty(ref sheet, value);
            ProcessFile.NotifyCanExecuteChanged();
        }
        get { return sheet; }
    }

    public FileResult? File
    {
        set
        {
            SetProperty(ref file, value);;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileLoaded)));
        }
        get { return file; }
    }

    public string FileName
    {
        get { return file != null ? $"Loaded File: '{file.FileName}'" : "No File Loaded"; }
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

    public bool ExportAdminTime
    {
        set { SetProperty(ref exportAdminTime, value); }
        get { return exportAdminTime; }
    }

    public bool FileLoaded
    {
        get { return !IsProcessing && File != null; }
    }


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

    [GeneratedRegex(@"[^\w\.]", RegexOptions.IgnoreCase, "en-US")]
    private static partial Regex CleanStringRegex();
}
