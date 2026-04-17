namespace StatsAnalyzer
{
    public class KillsList(SeasonInfo seasonInfo, String victim, String method, String killer)
    {
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;
        public String Victim { get; set; } = victim;
        public String Method { get; set; } = method;
        public String Killer { get; set; } = killer;
    }
}