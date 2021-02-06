using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using MeetMuch.Web.Calendar;
using MeetMuch.Web.Graph;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using TimeZoneConverter;

namespace MeetMuch.Web.Controllers
{
    public class TrendController : Controller
    {
        private readonly ICalendarAccess _calendarAccess;

        public TrendController(ICalendarAccess calendarAccess)
        {
            _calendarAccess = calendarAccess;
        }

        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
        public async Task<IActionResult> Index()
        {
            try
            {
                var start = DateTime.Today.AddDays(-14);
                var end = DateTime.Today;

                var events = await _calendarAccess.GetUserCalendar(User.GetUserGraphTimeZone(), start, end);

                // var meetingList = new ConcurrentDictionary<DateTime, Tuple<int, TimeSpan>>();
                // foreach (var calendarEvent in events)
                // {
                //     meetingList.AddOrUpdate(
                //         calendarEvent.Start.Date,
                //         new Tuple<int, TimeSpan>(1, calendarEvent.End - calendarEvent.Start),
                //         (time, old) => new Tuple<int, TimeSpan>(old.Item1 + 1, old.Item2 + (calendarEvent.End - calendarEvent.Start)));
                // }

                var result = events.Select(m => new TrendData
                {
                    Start = m.Start,
                    End = m.End
                });

                // Return a JSON dump of events
                return new ContentResult
                {
                    Content = JsonSerializer.Serialize(result.OrderBy(r => r.Start)),
                    ContentType = "application/json"
                };
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw;
                }

                return new ContentResult
                {
                    Content = $"Error getting calendar view: {ex.Message}",
                    ContentType = "text/plain"
                };
            }
        }

        private class TrendData
        {
            // public string Date { get; set; }
            // public int Minutes { get; set; }
            // public int Meetings { get; set; }
            public DateTime Start { get; set; }
            public DateTime End { get; set; }
        }
    }
}