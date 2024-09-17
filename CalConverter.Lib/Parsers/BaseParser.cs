
using CalConverter.Lib.Models;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Immutable;

namespace CalConverter.Lib;

public abstract class BaseParser
{
    /// <summary>
    /// <see cref="https://stackoverflow.com/a/4730266"/>
    /// </summary>
    public ImmutableDictionary<int, bool> IsNumberFormatDateMap { get; } = (new Dictionary<int, bool>() {
        { 0 , false }, //= 'General';
        { 1 , false }, //= '0';
        { 2 , false }, //= '0.00';
        { 3 , false }, //= '#,##0';
        { 4 , false }, //= '#,##0.00';
        { 9 , false }, //= '0%';
        { 10, false }, // = '0.00%';
        { 11, false }, // = '0.00E+00';
        { 12, false }, // = '# ?/?';
        { 13, false }, // = '# ??/??';
        { 14, true }, // = 'mm-dd-yy';
        { 15, true }, // = 'd-mmm-yy';
        { 16, true }, // = 'd-mmm';
        { 17, true }, // = 'mmm-yy';
        { 18, false }, // = 'h:mm AM/PM';
        { 19, false }, // = 'h:mm:ss AM/PM';
        { 20, false }, // = 'h:mm';
        { 21, false }, // = 'h:mm:ss';
        { 22, true }, // = 'm/d/yy h:mm';
        { 37, false }, // = '#,##0 ;(#,##0)';
        { 38, false }, // = '#,##0 ;[Red](#,##0)';
        { 39, false }, // = '#,##0.00;(#,##0.00)';
        { 40, false }, // = '#,##0.00;[Red](#,##0.00)';
        { 44, false }, // = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
        { 45, true }, // = 'mm:ss';
        { 46, true }, // = '[h]:mm:ss';
        { 47, true }, // = 'mmss.0';
        { 48, false }, // = '##0.0E+0';
        { 49, false }, // = '@';
        { 27, true }, // = '[$-404]e/m/d';
        { 30, true }, // = 'm/d/yy';
        { 36, true }, // = '[$-404]e/m/d';
        { 50, true }, // = '[$-404]e/m/d';
        { 57, true }, // = '[$-404]e/m/d';
        { 59, false }, // = 't0';
        { 60, false }, // = 't0.00';
        { 61, false }, // = 't#,##0';
        { 62, false }, // = 't#,##0.00';
        { 67, false }, // = 't0%';
        { 68, false }, // = 't0.00%';
        { 69, false }, // = 't# ?/?';
        { 70, false }, // = 't# ??/??';
    }).ToImmutableDictionary();

    protected Dictionary<int, DocumentFormat.OpenXml.Spreadsheet.Fill> patternRgbLookup { get; set; } = new();
    protected Dictionary<int, DocumentFormat.OpenXml.Spreadsheet.Fill> cellFormatLookup { get; set; } = new();

    protected Dictionary<uint, string> themeColorLookup { get; set; } = new();

    protected Dictionary<int, bool> isDateFormattingLookup { get; set; } = new();
    protected Dictionary<int, SharedStringItem> textLookup { get; set; } = new();
    protected Dictionary<string, string> mergedCells { get; set; } = new();
    protected SortedDictionary<string, List<string>> mergedCellGroups { get; set; } = new();

    public abstract string SheetName { get; }

    public IEnumerable<ScheduleBlock> ProcessFile(string fileName, string sheetName)
    {
        
        using SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false);

        if (TryGetWorkSheet(document, sheetName, out WorksheetPart workSheetPart))
        {
            SheetData? sheetData = workSheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

            ParseStyleParts(document);
            ParseTextStrings(document);
            ParseMergeCells(workSheetPart);

            return ProcessSheet(fileName, sheetName, sheetData); 
        }
        else
        {
            throw new InvalidOperationException("Can't find sheet part");
        }
    }


    public abstract IEnumerable<ScheduleBlock> ProcessSheet(string fileName, string sheetName, SheetData? sheetData);

    public void ParseMergeCells(WorksheetPart worksheetPart)
    {
        var sheetMergeCells = worksheetPart.Worksheet
            .Elements<MergeCells>()
            .Cast<MergeCells>()
            .SelectMany(q => q.AsEnumerable().Cast<MergeCell>());
        foreach (var mergeCell in sheetMergeCells)
        {
            var cellInGroup = Utils.GetCells(mergeCell.Reference.Value);
            var firstCell = cellInGroup.First();
            mergedCellGroups.Add(firstCell, cellInGroup);
            foreach (var cell in cellInGroup)
            {
                mergedCells[cell] = firstCell;
            }
        }
    }

    public void ParseStyleParts(SpreadsheetDocument document)
    {

        var styleParts = document.WorkbookPart?.WorkbookStylesPart;

        patternRgbLookup = styleParts.Stylesheet.Fills.ChildElements.Cast<DocumentFormat.OpenXml.Spreadsheet.Fill>()
            .Select((value, index) => (value, index))
            .ToDictionary(pair => pair.index, pair => pair.value);

        cellFormatLookup = styleParts.Stylesheet.CellFormats.ChildElements.Cast<CellFormat>()
            .Select((value, index) => (patternRgbLookup[(int)value.FillId.Value], index))
            .ToDictionary(pair => pair.index, pair => pair.Item1);

        isDateFormattingLookup = styleParts.Stylesheet.CellFormats.ChildElements.Cast<CellFormat>()
            .Select((value, index) => (IsNumberFormatDateMap[(int)value.NumberFormatId.Value], index))
            .ToDictionary(pair => pair.index, pair => pair.Item1);

        isDateFormattingLookup = styleParts.Stylesheet.CellFormats.ChildElements.Cast<CellFormat>()
           .Select((value, index) => (IsNumberFormatDateMap[(int)value.NumberFormatId.Value], index))
           .ToDictionary(pair => pair.index, pair => pair.Item1);

        themeColorLookup = document.WorkbookPart?.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements.Cast<Color2Type>()
            .Select((value, index) => (value, index))
            .ToDictionary(q => (uint)q.index,  q => {
                if (q.value.RgbColorModelHex != null) {
                    return q.value.RgbColorModelHex.Val.Value;
                }
                if (q.value.SystemColor != null)
                {
                    return q.value.SystemColor.LastColor.Value;
                }
                return string.Empty;
            });
    }

    public void ParseTextStrings(SpreadsheetDocument document)
    {
        textLookup = document.WorkbookPart?.SharedStringTablePart.SharedStringTable.ChildElements.Cast<SharedStringItem>()
              .Select((value, index) => (value, index))
              .ToDictionary(pair => pair.index, pair => pair.value);
    }


    private bool TryGetWorkSheet(SpreadsheetDocument document, string sheetName, out WorksheetPart workSheetPart)
    {
        workSheetPart = default;
        Sheet? sheet = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        if (sheet is not null)
        {
            if (sheet.Id is not null)
            {
                // The specified worksheet does not exist.
                workSheetPart = (WorksheetPart?)document.WorkbookPart?.GetPartById(sheet.Id);
                if (workSheetPart != null)
                {
                    return true;
                }
            }
        }

        return false;
    }


    public string GetColorType(DocumentFormat.OpenXml.Spreadsheet.ColorType ct)
    {
        string o = "";
        if (ct == null)
        {
            o += "NULL";
        }
        else
        {
            if (ct.Auto != null)
            {
                o += "System auto color";
            }

            if (ct.Rgb != null)
            {
                o += ct.Rgb.Value;
            }

            if (ct.Indexed != null)
            {
                o += $"Indexed color -> ${ct.Indexed.Value}";

                //IndexedColors ic = (IndexedColors)styles.Stylesheet.Colors.IndexedColors.ChildElements[(int)bgc.Indexed.Value];         
            }

            if (ct.Theme != null)
            {
                //o += $"Theme -> {ct.Theme.Value}";

                //Color2Type c2t = (Color2Type)sd.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)ct.Theme.Value];

                //o += $"RGB color model hex -> {c2t.RgbColorModelHex.Val}";

                o += themeColorLookup[ct.Theme.Value];
            }

            //if (ct.Tint != null)
            //{
            //    o += $"Tint value -> {ct.Tint.Value}";
            //}
        }
        return o;
    }
    
    public SimpleCellData? GetCellData(Cell? c)
    {
        if (c is null)
        {
            return null;
        }
        else
        {
            bool isDate = false;
            string? text = c?.CellValue?.Text;
            CellDataType datatype = CellDataType.String;

            var fill = cellFormatLookup[(int)c.StyleIndex.Value];
            if (!string.IsNullOrEmpty(text))
            {
                if (c.DataType != null)
                {
                    if (c.DataType == CellValues.SharedString)
                    {
                        text = textLookup[int.Parse(text)].InnerText;
                    }
                }

                isDate = isDateFormattingLookup[(int)c.StyleIndex.Value];
                if (isDate)
                {
                    datatype = CellDataType.Date;
                    text = DateTime.FromOADate(double.Parse(text)).ToShortDateString();
                }

                if (text.ToUpper() == "AM" || text.ToUpper() == "PM")
                {
                    datatype = CellDataType.TimeShift;
                }

            }
            else
            {
                datatype = CellDataType.Empty;
            }
            string color = GetColorType(fill.PatternFill.ForegroundColor);
            string colorName = string.Empty;
            if (!string.IsNullOrEmpty(color))
            {
                colorName = Utils.GetColorName(color);
            }
            return new SimpleCellData(c.CellReference, text ?? "NO_DATA", datatype, colorName);
        }
    }
}
