namespace StatsAnalyzer
{
    public class GlobalGaps(String player, int dayGap, String startRound, String startSeason, DateTime startDate, String endRound, String endSeason, DateTime endDate)
    {
        public String Player { get; set; } = player;
        public int DayGap { get; set; } = dayGap;
        public String StartRound { get; set; } = startRound;
        public String StartSeason { get; set; } = startSeason;
        public DateTime StartDate { get; set; } = startDate;
        public String EndRound { get; set; } = endRound;
        public String EndSeason { get; set; } = endSeason;
        public DateTime EndDate { get; set; } = endDate;
    }
}