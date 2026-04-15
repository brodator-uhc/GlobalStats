namespace StatsAnalyzer
{
    public class RoundList
    {
        public String Round { get; set; }
        public int TotalSeasons { get; set; }
        public int RosterSize { get; set; } = 0;
        public DateTime RoundDebut { get; set; }

        public RoundList(String round, int totalSeasons, DateTime roundDebut)
        {
            Round = round;
            TotalSeasons = totalSeasons;
            RoundDebut = roundDebut;
        }
    }
}