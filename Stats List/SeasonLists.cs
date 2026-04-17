using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class SeasonLists
    {
        public List<String> SeasonRoster { get; set; } = [];
        public List<String> SeasonDebutant { get; set; } = [];
        public List<String> SeasonTopFrag { get; set; } = [];
        public List<String> SeasonWinnerAlive { get; set; } = [];
        public List<String> SeasonWinnerDead { get; set; } = [];
        public List<String> SeasonAlive { get; set; } = [];
        public List<String> SeasonRunnerUp { get; set; } = [];
        public List<String> SeasonTeams { get; set; } = [];
    }
}