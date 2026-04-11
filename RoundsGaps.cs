namespace GlobalStats
{
    public class RoundsGaps
    {
        public List<String> Player { get; set; } = new List<String>();
        public List<int> DayGap { get; set; } = new List<int>();
        public List<String> Round { get; set; } = new List<String>();
        public List<String> StartSeason { get; set; } = new List<String>();
        public List<String> EndSeason { get; set; } = new List<String>();
        public List<DateTime> GapDate { get; set; } = new List<DateTime>();

        public void AddRoundGap(String player, int dayGap, String round, String startSeason, String endSeason, DateTime gapDate)
        {
            Player.Add(player);
            DayGap.Add(dayGap);
            Round.Add(round);
            StartSeason.Add(startSeason);
            EndSeason.Add(endSeason);
            GapDate.Add(gapDate);
        }
    }
}