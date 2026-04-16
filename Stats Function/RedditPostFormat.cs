namespace StatsAnalyzer
{
    public class RedditPostFormat
    {
        public static void FormatRunnerUps(RedditPosts redditPosts, SeasonInfo seasonInfo, List<String> seasonRunnerUps, List<String> seasonTeams, 
            List<String> seasonWinnerDead, Dictionary<String, int> killboard)
        {
            String runnerUpPost = "";
            seasonRunnerUps.Sort();
            foreach (String runnerUp in seasonRunnerUps)
            {
                if (killboard.ContainsKey(runnerUp))
                {
                    runnerUpPost += runnerUp + " (" + killboard[runnerUp] + "), ";
                }
                else
                {
                    runnerUpPost += runnerUp + " (0), ";
                }
            }
            if (seasonTeams.Count == 1 && !seasonWinnerDead.First().Equals("Ender Dragon"))
            {
                redditPosts.RunnerUps.Add("**S" + seasonInfo.SeasonNumber + ":** " + "N/A" + Environment.NewLine);
            }
            else
            {
                runnerUpPost = runnerUpPost[..^2];
                redditPosts.RunnerUps.Add("**S" + seasonInfo.SeasonNumber + ":** " + runnerUpPost + Environment.NewLine);
            }
        }
    }
}