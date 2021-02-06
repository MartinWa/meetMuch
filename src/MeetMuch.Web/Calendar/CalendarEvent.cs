using System;

namespace MeetMuch.Web.Calendar
{
    public class CalendarEvent
    {
        public string Subject { get; private set; }
        public string Organizer { get; private set; }
        public DateTime Start { get; internal set; }
        public DateTime End { get; internal set; }
    }
}