namespace StatsAnalyzer
{
    public class RostersList(SeasonInfo seasonInfo)
    {
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;
        public List<String> Roster { get; set; } = [];
    }
}