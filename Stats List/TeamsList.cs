namespace StatsAnalyzer
{
    public class TeamsList(SeasonInfo seasonInfo, String team)
    {
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;
        public String Team { get; set; } = team;
        public String TeamColor { get; set; } = "N/A";
    }
}