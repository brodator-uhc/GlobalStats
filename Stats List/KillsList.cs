namespace StatsAnalyzer
{
    public class KillsList(String round, String season, DateTime date, String victim, String method, String killer)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public String Victim { get; set; } = victim;
        public String Method { get; set; } = method;
        public String Killer { get; set; } = killer;
    }
}