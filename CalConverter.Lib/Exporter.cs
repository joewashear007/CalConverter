using CalConverter.Lib.Models;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib;
public class Exporter()
{
    public ExportOptions Options { get; set; } = new ExportOptions();

    public Dictionary<string, Calendar> Calendars { get; private set; } = [];


    public MemoryStream ToStream(string key = "ALL")
    {
        if (Calendars.TryGetValue(key, out var calendar))
        {
            MemoryStream stream = new();
            var serializer = new CalendarSerializer();
            serializer.Serialize(calendar, stream, Encoding.UTF8);
            return stream;
        } else
        {
            throw new ArgumentException("Can't fine the key");
        }
    }


    public void AddToCalendar(ScheduleBlock block)
    {
        TimeOnly MorningShiftStart = new TimeOnly(8, 0, 0);
        TimeOnly AfternoonShiftStart = new TimeOnly(13, 0, 0);

        if (DateOnly.TryParse(block.Date.Value, out var date))
        {
            if (Options.ExportStartDate <= date && date <= Options.ExportEndDate)
            {
                foreach (var preceptor in block.MorningShift.Percepters)
                {
                    CalendarEvent @event = CreateEventFromShift(MorningShiftStart, date, preceptor);
                    AddToCalendar(preceptor, @event);
                }
                foreach (var preceptor in block.MorningShift.Admins)
                {
                    CalendarEvent @event = CreateEventFromShift(MorningShiftStart, date, preceptor, isAdminTime: true);
                    AddToCalendar(preceptor, @event);
                }

                foreach (var preceptor in block.AfternoonShift.Percepters)
                {
                    CalendarEvent @event = CreateEventFromShift(AfternoonShiftStart, date, preceptor);
                    AddToCalendar(preceptor, @event);
                }
                foreach (var preceptor in block.AfternoonShift.Admins)
                {
                    CalendarEvent @event = CreateEventFromShift(AfternoonShiftStart, date, preceptor, isAdminTime: true);
                    AddToCalendar(preceptor, @event);
                }
            }
        }
        else
        {
            throw new InvalidOperationException($"Can't convert {date} to date");
        }
    }

    /// <summary>
    /// Add to calendar based on preceptor
    /// </summary>
    /// <param name="preceptor"></param>
    /// <param name="event"></param>
    private void AddToCalendar(ScheduleBlockPerson preceptor, CalendarEvent @event)
    {
        string key = Options.FilePerPerson ? preceptor.Attending.Value.ToLower() : "ALL";
        if (Calendars.TryGetValue(key, out var calendar))
        {
            calendar.Events.Add(@event);
        }
        else
        {
            calendar = new Calendar();
            calendar.Events.Add(@event);
            Calendars.Add(key, calendar);
        }
    }

    private CalendarEvent CreateEventFromShift(TimeOnly startTime, DateOnly date, ScheduleBlockPerson preceptor, bool isAdminTime = false)
    {
        var attendee = new Attendee
        {
            CommonName = preceptor.Attending.Value,
            Value = new Uri($"mailto:example@abc.com")
        };
        if (preceptor.Attending.Value != null && Options.UserEmailMap.TryGetValue(preceptor.Attending.Value, out string email))
        {
            attendee.Value = new Uri($"mailto:{email}");
        }
        var @event = new CalendarEvent()
        {
            Start = new CalDateTime(date.ToDateTime(startTime)),
            Duration = new TimeSpan(4, 0, 0),
            Attendees = [attendee]
        };
        if (isAdminTime)
        {
            @event.Resources.Add("Admin Time");
            @event.Summary = preceptor.Attending.Value + " [Admin Time]";
        }
        else
        {
            @event.Summary = preceptor.Attending.Value;
            if (preceptor.Room != null && !string.IsNullOrEmpty(preceptor.Room.Value))
            {
                @event.Resources.Add(preceptor.Room.Value);
            }
            if (!string.IsNullOrEmpty(preceptor.Attending.Color))
            {
                @event.Resources.Add(preceptor.Attending.Color + " Team");
                @event.Summary += " [" + preceptor.Attending.Color + "Team]";
            }

        }

        return @event;
    }
}
