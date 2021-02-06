using System;
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
                var userTimeZone = TZConvert.GetTimeZoneInfo(User.GetUserGraphTimeZone());
                var startOfWeek = _calendarAccess.GetUtcStartOfWeekInTimeZone(DateTime.Today, userTimeZone);

                var events = await _calendarAccess.GetUserWeekCalendar(startOfWeek, User.GetUserGraphTimeZone());

                // Return a JSON dump of events
                return new ContentResult
                {
                    Content = JsonSerializer.Serialize(events),
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
    }
}