namespace StatsAnalyzer
{
    public class GapRosterList(String round)
    {
        public String Round { get; set; } = round;
        public Dictionary<String, String> Roster { get; set; } = [];
    }
}