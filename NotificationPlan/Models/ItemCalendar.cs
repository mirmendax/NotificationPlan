using System;

namespace NotificationPlan.Models
{
    public class ItemCalendar
    {
        public string Body { get; set; }
        public string Title { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public int ReminderDay { get; set; }
        
        public bool IsReminded { get; set; }
        public int ReminderMinute  => ReminderDay * 24 * 60; 
        /// <summary>
        /// Заявка
        /// </summary>
        public bool IsRequest { get; set; }

        public override string ToString()
        {
            return $"{nameof(Body)}: {Body}, " +
                   $"{nameof(Title)}: {Title}, " +
                   $"{nameof(StartDateTime)}: {StartDateTime}, " +
                   $"{nameof(EndDateTime)}: {EndDateTime}, " +
                   $"{nameof(ReminderDay)}: {ReminderDay}, " +
                   $"{nameof(ReminderMinute)}: {ReminderMinute}, " +
                   $"{nameof(IsRequest)}: {IsRequest}";
        }
    }
}