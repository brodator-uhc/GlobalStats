namespace StatsAnalyzer
{
    public class TeamTypeList(String teamType, int timesUsed)
    {
        public String TeamType { get; set; } = teamType;
        public int TimesUsed { get; set; } = timesUsed;

        public static void UpdateTeamType(List<TeamTypeList> teamTypeLists, String teamType)
        {
            var teamTypeStats = teamTypeLists.Find(p => p.TeamType == teamType);
            teamTypeStats?.TimesUsed += 1;
        }
    }
}