using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib.Models;
public class ScheduleBlock
{
    public SimpleCellData Date { get; internal set; }
    public ScheduleBlockShift MorningShift { get; set; } = new() { ShiftBlock = ShiftBlock.AM };
    public ScheduleBlockShift AfternoonShift { get; set; } = new() { ShiftBlock = ShiftBlock.PM };

    public override string ToString()
    {
        return $@"
Schedule: {Date.Value}
AM: [{string.Join(", ", MorningShift.Percepters)}] / [{string.Join(",", MorningShift.Admins)}]
PM: [{string.Join(", ", AfternoonShift.Percepters)}] / [{string.Join(",", AfternoonShift.Admins)}]
";
    }
}

public enum ShiftBlock { AM, PM }
public class ScheduleBlockShift
{
    public ShiftBlock ShiftBlock { get; set; }
    public List<ScheduleBlockPerson> Percepters { get; set; } = new();
    public List<ScheduleBlockPerson> Admins { get; set; } = new();
}

public class ScheduleBlockPerson
{
    public SimpleCellData Room { get; set; }
    public SimpleCellData Attending { get; set; }

    public override string ToString()
    {
        if (string.IsNullOrEmpty(Room?.Value))
        {
            return $"{Attending?.Value}";
        }
        else
        {
            return $"{Room?.Value}|{Attending?.Value}({Attending?.Color})";
        }
    }
}