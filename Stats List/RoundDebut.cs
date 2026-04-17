namespace StatsAnalyzer
{
    public class RoundDebut(SeasonInfo seasonInfo, String player)
    {
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;
        public String Player { get; set; } = player;

        public static void UpdateRoundDebut(RoundDebut roundDebut, SeasonInfo seasonInfo)
        {
            roundDebut.Round = seasonInfo.SeasonName;
            roundDebut.Season = seasonInfo.SeasonNumber;
            roundDebut.Date = seasonInfo.SeasonDate;
        }
    }
}