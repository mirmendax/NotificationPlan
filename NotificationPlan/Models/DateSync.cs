namespace NotificationPlan.Models
{
    public class DateSync
    {
        public int Day { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public int Hour { get; set; }
        public int Minute { get; set; }

        public override string ToString()
        {
            return $"{Day}-{Month}-{Year} {Hour}:{Minute}";
        }
    }
}