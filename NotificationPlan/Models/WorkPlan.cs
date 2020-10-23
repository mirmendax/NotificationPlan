using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificationPlan.Models
{
    public class WorkPlan
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string ViewTO { get; set; }
        public DateTime StartTO { get; set; }
        public DateTime StopTO { get; set; }
        public string NameEmploy { get; set; }

        public bool IsDone { get; set; }

        public bool IsNotification { get; set; }

        public DateTime NotifyDate { get; set; }
    }
}
