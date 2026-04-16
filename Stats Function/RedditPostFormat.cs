namespace StatsAnalyzer
{
    public class RedditPostFormat
    {
        public static void FormatWins(RedditPosts redditPosts, SeasonInfo seasonInfo, List<String> seasonTeams, List<String> seasonWinnerAlive, 
            List<String> seasonWinnerDead, Dictionary<String, int> killboard, String seasonNumberPost)
        {
            String winnerPost = " ";
            seasonWinnerAlive.Sort();
            seasonWinnerDead.Sort();
            foreach (String winner in seasonWinnerAlive)
            {
                if (killboard.TryGetValue(winner, out int kills))
                {
                    winnerPost += winner + " (" + kills + "), ";
                }
                else
                {
                    winnerPost += winner + " (0), ";
                }
            }
            if (seasonWinnerDead.Count == 0)
            {
                winnerPost = winnerPost[..^2];
                winnerPost += "***";
            }
            else
            {
                if (!winnerPost.Equals(" "))
                {
                    winnerPost = winnerPost[..^2];
                    winnerPost += "**, *";
                }
                else
                {
                    winnerPost = "** *";
                }
            }
            foreach (String winner in seasonWinnerDead)
            {
                if (killboard.TryGetValue(winner, out int kills))
                {
                    winnerPost += winner + " (" + kills + "), ";
                }
                else
                {
                    if (winner.Equals("Ender Dragon"))
                    {
                        winnerPost += winner + ", ";
                    }
                    else
                    {
                        winnerPost += winner + " (0), ";
                    }
                }
            }
            winnerPost = winnerPost[..^2];
            winnerPost += "*";
            redditPosts.Winners.Add("**S" + seasonNumberPost + ":" + winnerPost + Environment.NewLine);
        }
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