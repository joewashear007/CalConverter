using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib.Models;

public enum CellDataType
{
    String,
    Number,
    Date,
    TimeShift,
    Empty
}
public record SimpleCellData(string CellRef, string Value, CellDataType DataType, string Color)
{
    public override string ToString()
    {
        return CellRef + ": " + Value + $"|{DataType}| [" + Color + "]";
    }
}
