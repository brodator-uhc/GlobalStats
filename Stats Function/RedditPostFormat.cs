namespace StatsAnalyzer
{
    public class RedditPostFormat
    {
        public static void FormatTopFrags(RedditPosts redditPosts, List<String> seasonTopFrag, int topFragAmount, String seasonNumberPost)
        {
            seasonTopFrag.Sort();
            String mostKillsPost = "";
            foreach (String topFrags in seasonTopFrag)
            {
                mostKillsPost = mostKillsPost + topFrags + ", ";
            }
            mostKillsPost = mostKillsPost[..^2];
            redditPosts.MostKills.Add("**S" + seasonNumberPost + ":** " + mostKillsPost + " (" + topFragAmount + ")" + Environment.NewLine);
        }
        public static void FormatTeamTopFrags(RedditPosts redditPosts, SeasonInfo seasonInfo, Dictionary<String, int> killboard,
            Dictionary<String, int> teamKillboard, String seasonNumberPost)
        {
            if (seasonInfo.IsFFA == false)
            {
                String mostKillsTeamPost = "";
                int teamTopFragsAmount = teamKillboard.Values.Max();
                foreach (String team in teamKillboard.Keys)
                {
                    if (teamKillboard[team] == teamTopFragsAmount)
                    {
                        String[] team_player = team.Split(',');

                        foreach (String player in team_player)
                        {
                            if (killboard.TryGetValue(player, out int kills))
                            {
                                mostKillsTeamPost += player + " (" + kills + "), ";
                            }
                            else
                            {
                                mostKillsTeamPost += player + " (0), ";
                            }
                        }

                        if (!mostKillsTeamPost.Equals(""))
                        {
                            mostKillsTeamPost = mostKillsTeamPost[..^2];
                            mostKillsTeamPost += " & ";
                        }
                    }
                }
                mostKillsTeamPost = mostKillsTeamPost[..^3];
                redditPosts.MostKillsTeam.Add("**S" + seasonNumberPost + ":** " + mostKillsTeamPost + Environment.NewLine);
            }
            else
            {
                redditPosts.MostKillsTeam.Add("**S" + seasonNumberPost + ":** " + "N/A" + Environment.NewLine);
            }
        }
        public static void FormatWins(RedditPosts redditPosts, List<String> seasonTeams, List<String> seasonWinnerAlive,
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
        public static void FormatRunnerUps(RedditPosts redditPosts, List<String> seasonRunnerUps, List<String> seasonTeams,
            List<String> seasonWinnerDead, Dictionary<String, int> killboard, String seasonNumberPost)
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
                redditPosts.RunnerUps.Add("**S" + seasonNumberPost + ":** " + "N/A" + Environment.NewLine);
            }
            else
            {
                runnerUpPost = runnerUpPost[..^2];
                redditPosts.RunnerUps.Add("**S" + seasonNumberPost + ":** " + runnerUpPost + Environment.NewLine);
            }
        }
    }
}