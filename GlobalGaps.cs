namespace GlobalStats
{
    public class GlobalGaps
    {
        public List<String> Player { get; set; } = new List<String>();
        public List<int> DayGap { get; set; } = new List<int>();
        public List<String> StartRound { get; set; } = new List<String>();
        public List<String> StartSeason { get; set; } = new List<String>();
        public List<DateTime> StartDate { get; set; } = new List<DateTime>();
        public List<String> EndRound { get; set; } = new List<String>();
        public List<String> EndSeason { get; set; } = new List<String>();
        public List<DateTime> EndDate { get; set; } = new List<DateTime>();

        public void AddGlobalGap(String player, int dayGap, String startRound, String startSeason, DateTime startDate, String endRound, String endSeason, DateTime endDate)
        {
            Player.Add(player);
            DayGap.Add(dayGap);
            StartRound.Add(startRound);
            StartSeason.Add(startSeason);
            StartDate.Add(startDate);
            EndRound.Add(endRound);
            EndSeason.Add(endSeason);
            EndDate.Add(endDate);
        }
    }
}