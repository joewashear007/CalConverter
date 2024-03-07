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
    public Dictionary<string, string> EmailMapping { get; set; } = new Dictionary<string, string>();

    public Calendar Calendar { get; private set; } =  new Calendar();


    public MemoryStream ToStream()
    {
        MemoryStream stream = new();
        var serializer = new CalendarSerializer();
        serializer.Serialize(Calendar, stream, Encoding.UTF8);
        return stream;
    }


    public void AddToCalendar(ScheduleBlock block)
    {
        TimeOnly MorningShiftStart = new TimeOnly(8, 0, 0);
        TimeOnly AfternoonShiftStart = new TimeOnly(13, 0, 0);

        if (DateOnly.TryParse(block.Date.Value, out var date))
        {
            foreach (var preceptor in block.MorningShift.Percepters)
            {
                CalendarEvent @event = CreateEventFromShift(MorningShiftStart, date, preceptor);
                Calendar.Events.Add(@event);    
            }
            foreach (var preceptor in block.MorningShift.Admins)
            {
                CalendarEvent @event = CreateEventFromShift(MorningShiftStart, date, preceptor, isAdminTime: true);
                Calendar.Events.Add(@event);
            }

            foreach (var preceptor in block.AfternoonShift.Percepters)
            {
                CalendarEvent @event = CreateEventFromShift(AfternoonShiftStart, date, preceptor);
                Calendar.Events.Add(@event);
            }
            foreach (var preceptor in block.AfternoonShift.Admins)
            {
                CalendarEvent @event = CreateEventFromShift(AfternoonShiftStart, date, preceptor, isAdminTime: true);
                Calendar.Events.Add(@event);
            }

        }
        else
        {
            throw new InvalidOperationException($"Can't convert {date} to date");
        }
    }

    private CalendarEvent CreateEventFromShift(TimeOnly startTime, DateOnly date, ScheduleBlockPerson preceptor, bool isAdminTime = false)
    {
        var attendee = new Attendee
        {
            CommonName = preceptor.Attending.Value,
            Value = new Uri($"mailto:example@abc.com")
        };    
        if (preceptor.Attending.Value != null && EmailMapping.TryGetValue(preceptor.Attending.Value, out string email))
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
