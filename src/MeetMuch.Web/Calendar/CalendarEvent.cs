using System;
using Microsoft.Graph;

namespace MeetMuch.Web.Calendar
{
    public class CalendarEvent
    {
        public string Subject { get; private set; }
        public string Organizer { get; private set; }
        public DateTime Start { get; private set; }
        public DateTime End { get; private set; }

        public CalendarEvent(Event graphEvent)
        {
            Subject = graphEvent.Subject;
            Organizer = graphEvent.Organizer.EmailAddress.Name;
            Start = DateTime.Parse(graphEvent.Start.DateTime);
            End = DateTime.Parse(graphEvent.End.DateTime);
        }
    }
}