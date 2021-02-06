using System;
using System.Threading.Tasks;
using MeetMuch.Web.Alerts;
using MeetMuch.Web.Calendar;
using MeetMuch.Web.Graph;
using MeetMuch.Web.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using TimeZoneConverter;

namespace MeetMuch.Web.Controllers
{
    public class CalendarController : Controller
    {
        private readonly ICalendarAccess _calendarAccess;

        public CalendarController(ICalendarAccess calendarAccess)
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
                var startOfWeekUtc = _calendarAccess.GetUtcStartOfWeekInTimeZone(DateTime.Today, userTimeZone);
                var events = await _calendarAccess.GetUserWeekCalendar(User.GetUserGraphTimeZone(), startOfWeekUtc);

                // Convert UTC start of week to user's time zone for
                // proper display
                var startOfWeekInTz = TimeZoneInfo.ConvertTimeFromUtc(startOfWeekUtc, userTimeZone);
                var model = new CalendarViewModel(startOfWeekInTz, events);

                return View(model);
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw;
                }

                return View(new CalendarViewModel())
                    .WithError("Error getting calendar view", ex.Message);
            }
        }
    }
}