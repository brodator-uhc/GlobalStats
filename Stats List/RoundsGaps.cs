namespace GlobalStats
{
    public class RoundsGaps
    {
        public String Player { get; set; }
        public int DayGap { get; set; }
        public String Round { get; set; }
        public String StartSeason { get; set; }
        public String EndSeason { get; set; }
        public DateTime GapDate { get; set; }

        public RoundsGaps(String player, int dayGap, String round, String startSeason, String endSeason, DateTime gapDate)
        {
            Player = player;
            DayGap = dayGap;
            Round = round;
            StartSeason = startSeason;
            EndSeason = endSeason;
            GapDate = gapDate;
        }
    }
}