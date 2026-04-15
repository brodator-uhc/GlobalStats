namespace StatsAnalyzer
{
    public class RostersList(String round, String season, DateTime date)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public List<String> Roster { get; set; } = [];
    }
}