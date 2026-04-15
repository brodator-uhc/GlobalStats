namespace StatsAnalyzer
{
    public class GlobalStats(String player)
    {
        public String Player { get; set; } = player;
        public int SeasonsPlayed { get; set; } = 0;
        public int Wins { get; set; } = 0;
        public int Alives { get; set; } = 0;
        public int RunnerUps { get; set; } = 0;
        public int Kills { get; set; } = 0;
        public int TopFrags { get; set; } = 0;
        public int PveDeaths { get; set; } = 0;
        public int FirstBloods { get; set; } = 0;
        public int FirstDeaths { get; set; } = 0;
        public int Ironmans { get; set; } = 0;
        public int FirstDamages { get; set; } = 0;
        public int Deaths { get; set; } = 0;
        public int TotalUniques { get; set; } = 0;
        public Double KDR { get; set; } = 0;
        public Double KPR { get; set; } = 0;

        public static void UpdateSeasonsPlayed(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.SeasonsPlayed += 1;
        }
        public static void UpdateWins(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.Wins += 1;
        }
        public static void UpdateAlives(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.Alives += 1;
        }
        public static void UpdateRunnerUps(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.RunnerUps += 1;
        }
        public static void UpdateKills(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.Kills += 1;
        }
        public static void UpdateTopFrags(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.TopFrags += 1;
        }
        public static void UpdatePveDeaths(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.PveDeaths += 1;
        }
        public static void UpdateFirstBloods(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.FirstBloods += 1;
        }
        public static void UpdateFirstDeaths(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.FirstDeaths += 1;
        }
        public static void UpdateIronmans(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.Ironmans += 1;
        }
        public static void UpdateFirstDamages(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.FirstDamages += 1;
        }
        public static void UpdateDeaths(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.Deaths += 1;
        }
        public static void UpdateTotalUniques(List<GlobalStats> globalStats, String player)
        {
            var playerGlobalStats = globalStats.Find(p => p.Player == player);
            playerGlobalStats?.TotalUniques += 1;
        }
        public static void UpdateKDRs(List<GlobalStats> globalStatsList)
        {
            foreach (GlobalStats globalStats in globalStatsList)
            {
                if (globalStats.Deaths.Equals(0))
                {
                    globalStats.KDR = Convert.ToDouble(globalStats.Kills);
                }
                else
                {
                    globalStats.KDR = Convert.ToDouble(globalStats.Kills) / Convert.ToDouble(globalStats.Deaths);
                }
                globalStats.KPR = Convert.ToDouble(globalStats.Kills) / Convert.ToDouble(globalStats.SeasonsPlayed);
            }
        }
    }
}