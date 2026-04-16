namespace StatsAnalyzer
{
    public class StatsList(SeasonInfo seasonInfo, String stat)
    {
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;
        public String Stat { get; set; } = stat;
    }
}