namespace StatsAnalyzer
{
    public class RoundDebut(String round, String season, DateTime date, String player)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public String Player { get; set; } = player;

        public static void UpdateRoundDebut(RoundDebut roundDebut, String round, String season, DateTime date)
        {
            roundDebut.Round = round;
            roundDebut.Season = season;
            roundDebut.Date = date;
        }
    }
}