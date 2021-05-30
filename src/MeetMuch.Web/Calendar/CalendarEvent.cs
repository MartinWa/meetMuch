using System;
using Microsoft.Graph;

namespace MeetMuch.Web.Calendar
{
    public class CalendarEvent
    {
        public string Subject { get; internal set; }
        public string Organizer { get; internal set; }
        public DateTime Start { get; internal set; }
        public DateTime End { get; internal set; }
        public bool? IsAllDay { get; internal set; }
        public ResponseType? Response { get; internal set; }
    }
}