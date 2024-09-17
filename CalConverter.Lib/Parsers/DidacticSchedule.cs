using CalConverter.Lib.Models;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CalConverter.Lib.Parsers;

public class DidacticSchedule : BaseParser
{
    public override string SheetName { get => "Lecture & Presenter List"; }

    public override IEnumerable<ScheduleBlock> ProcessSheet(string fileName, string sheetName, SheetData? sheetData)
    {
        List<ScheduleBlock> blocks = [];
        var cell = GetCellData(sheetData.Descendants<Cell>().FirstOrDefault(q => q.CellReference == "A2"));

        while (cell.Value != "NO_DATA")
        {
            var lecture1NameCell = sheetData.GetRelativeCell(cell, colOffset: 3);
            if (mergedCellGroups.ContainsKey(lecture1NameCell.CellReference))
            {
                // skip 
            }
            else
            {
                blocks.Add(FindSecheduled(sheetData, cell));
            }
            var cellRef = sheetData.GetRelativeCell(cell, rowOffset: 1);
            cell = GetCellData(cellRef);

        }

        return blocks.Where(q => q.AfternoonShift.Percepters.Any() || q.MorningShift.Percepters.Any());

    }


    private ScheduleBlock FindSecheduled(SheetData sheetData, SimpleCellData cell)
    {
        var cyleCell = sheetData.GetRelativeCell(cell, colOffset: 1);
        var groupCell = sheetData.GetRelativeCell(cell, colOffset: 2);
        var lecture1NameCell = sheetData.GetRelativeCell(cell, colOffset: 3);
        var lecture1PrecentorCell = sheetData.GetRelativeCell(cell, colOffset: 4);
        var lecture2NameCell = sheetData.GetRelativeCell(cell, colOffset: 5);
        var lecture2PrecentorCell = sheetData.GetRelativeCell(cell, colOffset: 6);
        var lecture3NameCell = sheetData.GetRelativeCell(cell, colOffset: 7);
        var lecture3PrecentorCell = sheetData.GetRelativeCell(cell, colOffset: 8);
        var lecture4NameCell = sheetData.GetRelativeCell(cell, colOffset: 9);
        var lecture4PrecentorCell = sheetData.GetRelativeCell(cell, colOffset: 10);

        string cycle = GetCellData(cyleCell).Value;
        if (mergedCells.ContainsKey(cyleCell.CellReference))
        {
            var parent_cell = sheetData.Descendants<Cell>().FirstOrDefault(q => q.CellReference == mergedCells[cyleCell.CellReference]);
            cycle = GetCellData(parent_cell).Value;
        }
        
        string group = "Group "+ GetCellData(groupCell).Value;

        var block = new ScheduleBlock()
        {
            Date = cell,
            MorningShift = new ScheduleBlockShift()
            {
                Percepters = [
                    new ScheduleBlockPerson() {
                        Attending = GetCellData(lecture1PrecentorCell),
                        EventLabel = GetCellData(lecture1NameCell).Value + $" - {cycle}, {group}",
                        Duration = TimeSpan.FromMinutes(30),
                        StartTime = new TimeOnly(8, 30)
                    },
                    new ScheduleBlockPerson() {
                        Attending = GetCellData(lecture2PrecentorCell),
                        EventLabel = GetCellData(lecture2NameCell).Value + $" - {cycle}, {group}",
                        Duration = TimeSpan.FromMinutes(30),
                        StartTime = new TimeOnly(9, 30)
                    },
                    new ScheduleBlockPerson() {
                        Attending = GetCellData(lecture3PrecentorCell),
                        EventLabel = GetCellData(lecture3NameCell).Value + $" - {cycle}, {group}",
                        Duration = TimeSpan.FromMinutes(30),
                        StartTime = new TimeOnly(10, 30)
                    },
                    new ScheduleBlockPerson() {
                        Attending = GetCellData(lecture4PrecentorCell),
                        EventLabel = GetCellData(lecture4NameCell).Value + $" - {cycle}, {group}",
                        Duration = TimeSpan.FromMinutes(30),
                        StartTime = new TimeOnly(11, 30)
                    },
                ],
                ShiftBlock = ShiftBlock.AM
            }
        };
        return block;
    }
}
