using System;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using MeetMuch.Web.Calendar;
using MeetMuch.Web.Graph;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace MeetMuch.Web.Controllers
{
    public class ListController : Controller
    {
        private readonly ICalendarAccess _calendarAccess;

        public ListController(ICalendarAccess calendarAccess)
        {
            _calendarAccess = calendarAccess;
        }

        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
        public async Task<IActionResult> Index(int start = 0)
        {
            DateTime endDate = default;
            DateTime startDate = default;
            if (start >= 0)
            {
                startDate = DateTime.Today;
                endDate = DateTime.Today.AddDays(start);
            }
            else
            {
                startDate = DateTime.Today.AddDays(start);
                endDate = DateTime.Today;
            }
            try
            {
                var events = await _calendarAccess.GetUserCalendar(User.GetUserGraphTimeZone(), startDate, endDate);

                // Return a JSON dump of events
                return new ContentResult
                {
                    Content = JsonSerializer.Serialize(events.OrderBy(r => r.Start)),
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