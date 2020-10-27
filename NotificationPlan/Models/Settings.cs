using System;

namespace NotificationPlan.Models
{
    public class Settings
    {
        public string NameEmploy { get; set; }
        public string PathToWorkPlan { get; set; }
        public int MonthAddedInCalendar { get; set; }
        public int YearAddedInCalendar { get; set; }
        public string StartFileWorkPlanWord { get; set; }
        public string EndFileWorkPlanWord { get; set; }
        public int DayRemineder { get; set; }
        public DateSync LastSync { get; set; }
    }
}