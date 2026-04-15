namespace StatsAnalyzer
{
    public class StatsList(String round, String season, DateTime date, String stat)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public String Stat { get; set; } = stat;
    }
}