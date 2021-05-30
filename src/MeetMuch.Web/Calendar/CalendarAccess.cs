using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MeetMuch.Web.Calendar
{
    public interface ICalendarAccess
    {
        Task<IList<Event>> GetUserWeekCalendar(string userTimeZone, DateTime startOfWeekUtc);
        Task<IEnumerable<CalendarEvent>> GetUserCalendar(string userTimeZone, DateTime start, DateTime end);
        DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone);
    }

    public class CalendarAccess : ICalendarAccess
    {
        private readonly GraphServiceClient _graphClient;

        public CalendarAccess(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<IList<Event>> GetUserWeekCalendar(string userTimeZone, DateTime startOfWeekUtc)
        {
            // Configure a calendar view for the current week
            var endOfWeekUtc = startOfWeekUtc.AddDays(7);
            return await GetEvents(userTimeZone, startOfWeekUtc, endOfWeekUtc, e => new
            {
                e.Subject,
                e.Organizer,
                e.Start,
                e.End
            });
        }

        public async Task<IEnumerable<CalendarEvent>> GetUserCalendar(string userTimeZone, DateTime start, DateTime end)
        {
            var graphEvents = await GetEvents(userTimeZone, start.ToUniversalTime(), end.ToUniversalTime(), e => new
            {
                e.Start,
                e.End,
                e.Organizer,
                e.Subject,
                e.IsAllDay,
                e.ResponseStatus
            });
            return graphEvents == null
                ? Enumerable.Empty<CalendarEvent>()
                : graphEvents.Select(e => new CalendarEvent
                {
                    Start = DateTime.Parse(e.Start.DateTime),
                    End = DateTime.Parse(e.End.DateTime),
                    Subject = e.Subject,
                    Organizer = e.Organizer.EmailAddress.Name,
                    IsAllDay = e.IsAllDay,
                    Response = e.ResponseStatus.Response
                });
        }

        private async Task<IList<Event>> GetEvents(string userTimeZone, DateTime start, DateTime end, Expression<Func<Event, object>> selectExpression)
        {
            var viewOptions = new List<QueryOption>
            {
                new("startDateTime", start.ToString("o")),
                new("endDateTime", end.ToString("o"))
            };

            var events = await _graphClient.Me.CalendarView
                .Request(viewOptions)
                // Send user time zone in request so date/time in
                // response will be in preferred time zone
                .Header("Prefer", $"outlook.timezone=\"{userTimeZone}\"")
                // Get max 50 per request
                .Top(50)
                // Only return fields app will use
                .Select(selectExpression)
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