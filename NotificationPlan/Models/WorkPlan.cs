using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificationPlan.Models
{
    public class WorkPlan
    {
        public string Title { get; set; }
        public string ViewTO { get; set; }
        public DateTime StartTO { get; set; }
        public DateTime EndTO { get; set; }
        public string NameEmploy { get; set; }


        public override string ToString()
        {
            return $"{nameof(Title)}: {Title}, " +
                   $"{nameof(ViewTO)}: {ViewTO}, " +
                   $"{nameof(StartTO)}: {StartTO}, " +
                   $"{nameof(EndTO)}: {EndTO}, " +
                   $"{nameof(NameEmploy)}: {NameEmploy}, ";
        }
    }
}
