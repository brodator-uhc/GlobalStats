namespace StatsAnalyzer
{
    public class StatsList
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
        public String Stat { get; set; }

        public StatsList(String round, String season, DateTime date, String stat)
        {
            Round = round;
            Season = season;
            Date = date;
            Stat = stat;
        }
    }
}