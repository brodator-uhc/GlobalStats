namespace StatsAnalyzer
{
    public class RoundGapSheet(String player, String join, int number)
    {
        public int Number { get; set; } = number;
        public String Player { get; set; } = player;
        public String SeasonJoined { get; set; } = join;
        public int TotalPlayed { get; set; } = 0;
        public List<String> Participations { get; set; } = [];
    }
}