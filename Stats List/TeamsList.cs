namespace StatsAnalyzer
{
    public class TeamsList(String round, String season, DateTime date, String team)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public String Team { get; set; } = team;
        public String TeamColor { get; set; } = "N/A";
    }
}