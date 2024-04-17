using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib.Models;
public class ExportOptions
{
    public Dictionary<string, string> UserEmailMap { get; set; } = new();

    public DateOnly ExportStartDate { get; set; } = DateOnly.MinValue;

    public DateOnly ExportEndDate { get; set; } = DateOnly.MaxValue;

    public bool FilePerPerson { get; set; } = true;
}
