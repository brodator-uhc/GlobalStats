namespace StatsAnalyzer
{
    public class RoundList(String round, int totalSeasons, DateTime roundDebut)
    {
        public String Round { get; set; } = round;
        public int TotalSeasons { get; set; } = totalSeasons;
        public int RosterSize { get; set; } = 0;
        public DateTime RoundDebut { get; set; } = roundDebut;
    }
}