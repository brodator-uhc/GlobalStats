namespace StatsAnalyzer
{
    public class RoundsGaps(String player, int dayGap, String round, String startSeason, String endSeason, DateTime gapDate)
    {
        public String Player { get; set; } = player;
        public int DayGap { get; set; } = dayGap;
        public String Round { get; set; } = round;
        public String StartSeason { get; set; } = startSeason;
        public String EndSeason { get; set; } = endSeason;
        public DateTime GapDate { get; set; } = gapDate;
    }
}