using Ical.Net;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib;
public static class CalendarUtils
{
    public static string Summarize(Calendar calendar)
    {
        var events = calendar.Events
            .GroupBy(q => q.Start.Date)
            .OrderBy(q => q.Key)
            .ToList();


        StringBuilder summary = new StringBuilder();
        foreach (var dayGroup in events)
        {
            summary
                .AppendLine(dayGroup.Key.ToShortDateString())
                .AppendLine("---------------------------");
            var dayEvents = dayGroup.OrderByDescending(q => q.Start.Hour + q.Attendees.FirstOrDefault()?.CommonName ?? "").ToList();
            foreach (var dayEvent in dayEvents)
            {
                summary.Append(" * ").AppendFormat("{0:00}", dayEvent.Start.Hour).Append(":00 - ").Append(dayEvent.Summary);
                if (dayEvent.Resources.Count > 0)
                {
                    summary.Append(" ( Room: ").Append(dayEvent.Resources.First()).Append(")");
                }
                summary.AppendLine();
            }
            summary.AppendLine().AppendLine();
        }

        return summary.ToString();
    }
}
