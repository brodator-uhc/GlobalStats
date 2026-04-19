using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class RedditPostFormat
    {
        public static void FormatParticipations(RedditPosts redditPosts, SeasonLists seasonLists, String seasonNumberPost,
            int firstDataColumn, IXLCell seasonNumberCell)
        {
            String seasonDebutantPost = "";
            foreach (String debutant in seasonLists.SeasonDebutant)
            {
                seasonDebutantPost += debutant + ", ";
            }
            if (seasonDebutantPost.Length > 0)
            {
                seasonDebutantPost = seasonDebutantPost[..^2];
                redditPosts.Debutants.Add("**S" + seasonNumberPost + " (" + (seasonDebutantPost.Count(c => c == ',') + 1) + "):** " + seasonDebutantPost + Environment.NewLine);
            }
            else
            {
                redditPosts.Debutants.Add("**S" + seasonNumberPost + " (" + seasonDebutantPost.Count(c => c == ',') + "):** " + Environment.NewLine);
            }

            foreach (String player in seasonLists.SeasonRoster)
            {
                String lastSeason = "";
                if (firstDataColumn > 2)
                {
                    lastSeason = seasonNumberCell.CellLeft(3).GetString();
                }
                else
                {
                    lastSeason = "N/A";
                }
                RedditPostsPlayed.UpdateParticipations(redditPosts, player, seasonNumberPost, lastSeason);
            }
        }
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
        public static void FormatWins(RedditPosts redditPosts, SeasonInfo seasonInfo, WinnerInfo winnerInfo, SeasonLists seasonLists,
        Dictionary<String, int> killboard, String seasonNumberPost)
        {
            String winnerPost = " ";
            seasonLists.SeasonWinnerAlive.Sort();
            seasonLists.SeasonWinnerDead.Sort();
            foreach (String winner in seasonLists.SeasonWinnerAlive)
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
            if (seasonLists.SeasonWinnerDead.Count == 0)
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
            if (winnerInfo.IsDoubleKillWinner == false)
            {
                winnerPost = FormatDeadWinners(seasonLists.SeasonWinnerDead, killboard, winnerPost);
            }
            else
            {
                if (seasonInfo.IsFFA == false)
                {
                String firstWinningTeam = "";
                String secondWinningTeam = "";
                foreach (String team in seasonLists.SeasonTeams)
                {
                    foreach (String winner in seasonLists.SeasonWinnerDead)
                    {
                        if (team.Contains(winner))
                        {
                            if (firstWinningTeam == "")
                            {
                                firstWinningTeam = team;
                            } else
                            {
                                secondWinningTeam = team;
                            }
                        }
                    }
                }
                String[] firstTeamList = firstWinningTeam.Split(',');
                String[] secondTeamList = secondWinningTeam.Split(',');

                winnerPost = FormatDeadWinners([.. firstTeamList], killboard, winnerPost);
                winnerPost = winnerPost[..^2];
                winnerPost += " & ";
                winnerPost = FormatDeadWinners([.. secondTeamList], killboard, winnerPost);
                }
                else
                {
                    String[] firstWinner = [""];
                    String[] secondWinner = [""];

                    firstWinner[0] = seasonLists.SeasonWinnerDead[0];
                    secondWinner[0] = seasonLists.SeasonWinnerDead[1];

                    winnerPost = FormatDeadWinners([.. firstWinner], killboard, winnerPost);
                    winnerPost = winnerPost[..^2];
                    winnerPost += " & ";
                    winnerPost = FormatDeadWinners([.. secondWinner], killboard, winnerPost);
                }
            }
            winnerPost = winnerPost[..^2];
            winnerPost += "*";
            redditPosts.Winners.Add("**S" + seasonNumberPost + ":" + winnerPost + Environment.NewLine);
        }
        public static String FormatDeadWinners(List<String> winnerList, Dictionary<String, int> killboard, String winnerPost)
        {
            foreach (String winner in winnerList)
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
            return winnerPost;
        }
        public static void FormatRunnerUps(RedditPosts redditPosts, SeasonLists seasonLists, Dictionary<String, int> killboard, String seasonNumberPost)
        {
            String runnerUpPost = "";
            seasonLists.SeasonRunnerUp.Sort();
            foreach (String runnerUp in seasonLists.SeasonRunnerUp)
            {
                if (killboard.TryGetValue(runnerUp, out int kills))
                {
                    runnerUpPost += runnerUp + " (" + kills + "), ";
                }
                else
                {
                    runnerUpPost += runnerUp + " (0), ";
                }
            }
            if (seasonLists.SeasonTeams.Count == 1 && !seasonLists.SeasonWinnerDead.First().Equals("Ender Dragon"))
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