namespace StatsAnalyzer
{
    public class GlobalGaps
    {
        public String Player { get; set; }
        public int DayGap { get; set; }
        public String StartRound { get; set; }
        public String StartSeason { get; set; }
        public DateTime StartDate { get; set; }
        public String EndRound { get; set; }
        public String EndSeason { get; set; }
        public DateTime EndDate { get; set; }

        public GlobalGaps(String player, int dayGap, String startRound, String startSeason, DateTime startDate, String endRound, String endSeason, DateTime endDate)
        {
            Player = player;
            DayGap = dayGap;
            StartRound = startRound;
            StartSeason = startSeason;
            StartDate = startDate;
            EndRound = endRound;
            EndSeason = endSeason;
            EndDate = endDate;
        }
    }
}