namespace GlobalStats
{
    public class RoundDebut
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
        public String Player { get; set; }

        public RoundDebut(String round, String season, DateTime date, String player)
        {
            Round = round;
            Season = season;
            Date = date;
            Player = player;
        }

        public static void UpdateRoundDebut(RoundDebut roundDebut, String round, String season, DateTime date)
        {
            roundDebut.Round = round;
            roundDebut.Season = season;
            roundDebut.Date = date;
        }
    }
}