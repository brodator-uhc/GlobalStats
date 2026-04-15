namespace StatsAnalyzer
{
    public class TeamsList
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
        public String Team { get; set; }
        public String TeamColor { get; set; } = "N/A";

        public TeamsList(String round, String season, DateTime date, String team)
        {
            Round = round;
            Season = season;
            Date = date;
            Team = team;
        }
    }
}