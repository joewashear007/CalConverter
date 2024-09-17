using CalConverter.Lib.Models;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Reflection.Metadata;

namespace CalConverter.Lib.Parsers;

public class PreceptorSchedule : BaseParser
{
    public override string SheetName { get => "Sheet1"; }

    public override IEnumerable<ScheduleBlock> ProcessSheet(string fileName, string sheetName, SheetData? sheetData)
    {
        List<ScheduleBlock> blocks = [];
        var firstCell = GetCellData(sheetData.Descendants<Cell>().First());

        foreach (var cell in mergedCellGroups.Keys)
        {
            bool isDate = false;
            var cellObj = sheetData.Descendants<Cell>().FirstOrDefault(q => q.CellReference == cell);
            var simpleCell = GetCellData(cellObj);

            if (simpleCell.DataType == CellDataType.Date
                && simpleCell.Color != firstCell.Color)
            {
                blocks.Add(FindSecheduled(sheetData, simpleCell, firstCell));
            }
        }

        return blocks.Where(q => q.AfternoonShift.Percepters.Any() || q.MorningShift.Percepters.Any());

    }


    private SimpleCellData? GetCellData(Cell? c)
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


    private ScheduleBlock FindSecheduled(SheetData sheetData, SimpleCellData cell, SimpleCellData firstCell)
    {
        ScheduleBlock block = new() { Date = cell };
        TimeOnly MorningShiftStart = new TimeOnly(8, 0, 0);
        TimeOnly AfternoonShiftStart = new TimeOnly(13, 0, 0);

        TimeOnly startTime = MorningShiftStart;
        SimpleCellData neighborCell;
        SimpleCellData curCell = cell;
        ScheduleBlockShift curShift = block.MorningShift;
        List<ScheduleBlockPerson> curPersonList = curShift.Percepters;
        do
        {
            curCell = GetCellData(sheetData.GetRelativeCell(curCell, rowOffset: 1));
            if (curCell is not null)
            {
                if (curCell.DataType == CellDataType.TimeShift && curCell.Value == "PM")
                {
                    curShift = block.AfternoonShift;
                    curPersonList = curShift.Percepters;
                    startTime = AfternoonShiftStart;
                }
                else
                {
                    neighborCell = GetCellData(sheetData.GetRelativeCell(curCell, colOffset: 1));
                    if (curCell.DataType == CellDataType.Empty)
                    {
                        if (neighborCell.DataType == CellDataType.Empty)
                        {
                            if (curShift.Percepters.Count > 0)
                            {
                                // No Room and no Attending, break in the list person list if there some attendings
                                // trying to miss the stupid extra lines
                                curPersonList = curShift.Admins;
                            }
                        }
                        else
                        {
                            curPersonList.Add(new ScheduleBlockPerson() { Attending = neighborCell, StartTime = startTime });
                        }
                    }
                    if (curCell.DataType == CellDataType.String && neighborCell.DataType == CellDataType.String)
                    {
                        // special case ACUTE clinic, attending name is on the next line
                        if (neighborCell.Value.Trim().Equals("ACUTE", StringComparison.InvariantCultureIgnoreCase))
                        {
                            var nextCell = GetCellData(sheetData.GetRelativeCell(neighborCell, rowOffset: 1));
                            if (nextCell.DataType == CellDataType.String)
                            {
                                curPersonList.Add(new ScheduleBlockPerson() { Attending = nextCell, StartTime = startTime, Room = curCell, EventLabel = "Acute Clinic" });
                            }
                            // we are parsing two rows, so correct that in the increment loop
                            curCell = GetCellData(sheetData.GetRelativeCell(curCell, rowOffset: 1));
                            continue;
                        }
                        else
                        {
                            // a room / attending call
                            curPersonList.Add(new ScheduleBlockPerson() { Attending = neighborCell, StartTime = startTime, Room = curCell });
                        }
                    }
                    if (curCell.DataType == CellDataType.String && neighborCell.DataType == CellDataType.Empty)
                    {
                        // a room / no attending - what to do?
                        curPersonList.Add(new ScheduleBlockPerson() { Attending = neighborCell, StartTime = startTime, Room = curCell });
                    }
                }

            }
            else
            {
                break;
            }
        } while (curCell.Color != firstCell.Color);
        return block;
    }
}
