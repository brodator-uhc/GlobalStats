namespace StatsAnalyzer
{
    public class KillsList
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
        public String Victim { get; set; }
        public String Method { get; set; }
        public String Killer { get; set; }

        public KillsList(String round, String season, DateTime date, String victim, String method, String killer)
        {
            Round = round;
            Season = season;
            Date = date;
            Victim = victim;
            Method = method;
            Killer = killer;
        }
    }
}