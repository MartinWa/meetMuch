using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MeetMuch.Web.Calendar
{
    public interface ICalendarAccess
    {
        Task<IList<Event>> GetUserWeekCalendar(DateTime startOfWeekUtc, string userTimeZone);
        DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone);
    }

    public class CalendarAccess : ICalendarAccess
    {
        private readonly GraphServiceClient _graphClient;

        public CalendarAccess(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<IList<Event>> GetUserWeekCalendar(DateTime startOfWeekUtc, string userTimeZone)
        {
            // Configure a calendar view for the current week
            var endOfWeekUtc = startOfWeekUtc.AddDays(7);

            var viewOptions = new List<QueryOption>
            {
                new("startDateTime", startOfWeekUtc.ToString("o")),
                new("endDateTime", endOfWeekUtc.ToString("o"))
            };

            var events = await _graphClient.Me.CalendarView
                .Request(viewOptions)
                // Send user time zone in request so date/time in
                // response will be in preferred time zone
                .Header("Prefer", $"outlook.timezone=\"{userTimeZone}\"")
                // Get max 50 per request
                .Top(50)
                // Only return fields app will use
                .Select(e => new
                {
                    e.Subject,
                    e.Organizer,
                    e.Start,
                    e.End
                })
                // Order results chronologically
                .OrderBy("start/dateTime")
                .GetAsync();

            IList<Event> allEvents;
            // Handle case where there are more than 50
            if (events.NextPageRequest != null)
            {
                allEvents = new List<Event>();
                // Create a page iterator to iterate over subsequent pages
                // of results. Build a list from the results
                var pageIterator = PageIterator<Event>.CreatePageIterator(
                    _graphClient, events,
                    (e) =>
                    {
                        allEvents.Add(e);
                        return true;
                    }
                );
                await pageIterator.IterateAsync();
            }
            else
            {
                // If only one page, just use the result
                allEvents = events.CurrentPage;
            }

            return allEvents;
        }

        public DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone)
        {
            // Assumes Sunday as first day of week
            var diff = System.DayOfWeek.Sunday - today.DayOfWeek;

            // create date as unspecified kind
            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            // convert to UTC
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, timeZone);
        }
    }
}