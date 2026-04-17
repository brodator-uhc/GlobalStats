using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class PlayerStatsCalculator
    {
        public static void CalculatePlayerStats(List<PlayerStats> playerStatsList, IXLWorksheet rosterList, IXLWorkbook statsDocument, String playerStats)
        {
            //Makes the list of all rounds played, and sets the other variables
            for (int round = 1; round <= rosterList.Rows().Count(); round++)
            {
                IXLRange rosterRange = rosterList.Range(round, 4, round, 129);
                foreach (IXLCell cell in rosterRange.CellsUsed())
                {
                    string value = cell.GetString();

                    if (value == playerStats)
                    {
                        String roundName = rosterList.Cell(round, 1).Value.ToString();
                        String seasonNumber = rosterList.Cell(round, 2).Value.ToString();
                        DateTime seasonDate = rosterList.Cell(round, 3).Value.GetDateTime();
                        playerStatsList.Add(new PlayerStats(roundName, seasonNumber, seasonDate));
                    }
                }
            }
            Console.WriteLine("Roster Done");

            //Makes the list of the teams and formats it correctly
            //Also gets team color and makes it match the format used.
            var statsList = statsDocument.Worksheet(2);
            IXLRange teamsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell team in teamsRange.Cells())
            {
                if (team.Value.ToString().Contains(playerStats))
                {
                    //Gets team for the season
                    String roundName = team.CellLeft(3).Value.ToString();
                    String seasonNumber = team.CellLeft(2).Value.ToString();
                    String seasonTeam = team.Value.ToString() + ",";
                    seasonTeam = seasonTeam.Replace(playerStats + ",", "");
                    seasonTeam = seasonTeam.Replace(",", " ");

                    //Gets team color for the season
                    String seasonTeamColor = team.CellRight().Value.ToString() + "";
                    seasonTeamColor = PlayerStats.GetTeamColorChar(seasonTeamColor);

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    if (seasonStats != null)
                    {
                        seasonStats.Team = seasonTeam;
                        seasonStats.TeamColor = seasonTeamColor;
                    }
                }
            }
            Console.WriteLine("Teams Done");

            //Makes the list for all the kills and deaths of a player, formats it for the doc
            statsList = statsDocument.Worksheet(1);
            IXLRange killsRange = statsList.Range(1, 6, statsList.RangeUsed()!.RowCount(), 6);
            IXLRange deathsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell playerDeath in killsRange.CellsUsed())
            {
                if (playerDeath.Value.ToString() == playerStats)
                {
                    String roundName = playerDeath.CellLeft(5).Value.ToString();
                    String seasonNumber = playerDeath.CellLeft(4).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    if (seasonStats != null)
                    {
                        String playerKilled = playerDeath.CellLeft(2).Value.ToString();
                        if (seasonStats.Kills == "/")
                        {
                            seasonStats.Kills = playerKilled;
                        }
                        else
                        {
                            seasonStats.Kills = seasonStats.Kills + " " + playerKilled;
                        }
                    }
                }
            }
            Console.WriteLine("Kills Done");

            foreach (IXLCell playerDeath in deathsRange.CellsUsed())
            {
                if (playerDeath.Value.ToString() == playerStats)
                {
                    String roundName = playerDeath.CellLeft(3).Value.ToString();
                    String seasonNumber = playerDeath.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    if (seasonStats != null)
                    {
                        String playerDied = playerDeath.CellRight(2).Value.ToString();
                        if (seasonStats.Death == "todelete")
                        {
                            seasonStats.Death = playerDied;
                        }
                        else
                        {
                            seasonStats.Death = seasonStats.Death + " " + playerDied;
                        }
                    }
                }
            }
            Console.WriteLine("Deaths Done");

            //Calculates the number of kills in 1 season
            foreach (var round in playerStatsList.ToList())
            {
                if (round.Kills != "/")
                {
                    int count = round.Kills.Split(' ').Length;
                    round.KillsTotal = count;
                }
            }
            Console.WriteLine("Kill Count Done");

            //Gets the list for first damages
            statsList = statsDocument.Worksheet(3);
            IXLRange firstDamageRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in firstDamageRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.FirstDamage = "x";
                }
            }
            Console.WriteLine("First Damage Done");

            //Gets the list for ironman
            statsList = statsDocument.Worksheet(4);
            IXLRange ironmanRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in ironmanRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.Ironman = "x";
                }
            }
            Console.WriteLine("Ironman Done");

            //Gets the list for pve deaths
            statsList = statsDocument.Worksheet(5);
            IXLRange pveRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in pveRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.PveDeath = "x";
                }
            }
            Console.WriteLine("PvE Deaths Done");

            //Gets the list for first deaths
            statsList = statsDocument.Worksheet(6);
            IXLRange firstDeathsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in firstDeathsRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.FirstDeath = "x";
                }
            }
            Console.WriteLine("First Deaths Done");

            //Gets the list for first blood
            statsList = statsDocument.Worksheet(7);
            IXLRange firstBloodRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in firstBloodRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.FirstBlood = "x";
                }
            }
            Console.WriteLine("First Blood Done");

            //Gets the list for most kills
            statsList = statsDocument.Worksheet(8);
            IXLRange topFragsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in topFragsRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.TopFrag = "x";
                }
            }
            Console.WriteLine("Top Frags Done");

            //Gets the list for runner ups
            statsList = statsDocument.Worksheet(9);
            IXLRange runnerUpsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in runnerUpsRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.RunnerUp = "x";
                }
            }
            Console.WriteLine("Runner Ups Done");

            //Gets the list for wins
            statsList = statsDocument.Worksheet(11);
            IXLRange winRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in winRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    seasonStats?.Win = "x";
                }
            }
            Console.WriteLine("Wins Done");

            //Gets the list for alives wins, if not a win (dragon rush) logs in console
            statsList = statsDocument.Worksheet(10);
            IXLRange aliveRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

            foreach (IXLCell player in aliveRange.Cells())
            {
                if (player.Value.ToString().Equals(playerStats))
                {
                    String roundName = player.CellLeft(3).Value.ToString();
                    String seasonNumber = player.CellLeft(2).Value.ToString();

                    var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
                    if (seasonStats != null)
                    {
                        if (seasonStats.Win == "x")
                        {
                            seasonStats.Win = "o";
                        }
                        else
                        {
                            Console.WriteLine("Alive no win found!");
                        }
                    }
                }
            }
            Console.WriteLine("Alives Done");
        }
    }
}