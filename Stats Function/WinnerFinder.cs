using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class WinnerFinder
    {
        public static void SetSoloWinners(List<GlobalStats> globalStatsList, List<StatsList> winList, SeasonInfo seasonInfo, String winner,
            SeasonLists seasonLists)
        {
            if (seasonLists.SeasonAlive.Contains(winner))
            {
                seasonLists.SeasonWinnerAlive.Add(winner);
            }
            else
            {
                seasonLists.SeasonWinnerDead.Add(winner);
            }

            if (seasonInfo.IsCrossoverSeason == false)
            {
                GlobalStats.UpdateWins(globalStatsList, winner);
                winList.Add(new StatsList(seasonInfo, winner));
            }
        }
        public static void SetTeamWinners(List<GlobalStats> globalStatsList, List<StatsList> winList, SeasonInfo seasonInfo, WinnerInfo winnerInfo, 
            String teamWinner, SeasonLists seasonLists)
        {
            foreach (String team in seasonLists.SeasonTeams)
            {
                if (team.Contains(teamWinner))
                {
                    winnerInfo.WinningTeam = team;

                    //Splits the team string to get each player and gives them a win
                    String[] winnerSplit = team.Split(',');
                    foreach (String winner in winnerSplit)
                    {
                        if (seasonLists.SeasonAlive.Contains(winner))
                        {
                            seasonLists.SeasonWinnerAlive.Add(winner);
                        }
                        else
                        {
                            seasonLists.SeasonWinnerDead.Add(winner);
                        }

                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            GlobalStats.UpdateWins(globalStatsList, winner);
                            winList.Add(new StatsList(seasonInfo, winner));
                        }
                    }
                }
            }
        }
        public static void GetWinners(List<GlobalStats> globalStatsList, List<StatsList> winList, SeasonInfo seasonInfo, WinnerInfo winnerInfo,
            SeasonLists seasonLists)
        {
            if (winnerInfo.LastDataRowCell.CellRight(2).GetString().Equals("Nothing"))
            {
                //Gets the season winner on the last row
                String winner = winnerInfo.LastDataRowCell.GetString();
                winnerInfo.WinnerCell = winnerInfo.LastDataRowCell;

                //If gamemode with unique win condition
                if (seasonInfo.SeasonGamemodes.Contains("Dragon Rush") ||
                    seasonInfo.SeasonGamemodes.Contains("Wither Rush") ||
                    seasonInfo.SeasonGamemodes.Contains("Realm Rush") ||
                    seasonInfo.SeasonGamemodes.Contains("Bolas Rush") ||
                    seasonInfo.SeasonGamemodes.Contains("Escape From Gaia") ||
                    seasonInfo.SeasonGamemodes.Contains("Trouble In Paradise") ||
                    seasonInfo.SeasonGamemodes.Contains("Dragon Rush Deviation Version") ||
                    seasonInfo.SeasonGamemodes.Contains("Mega Spud Rush") ||
                    seasonInfo.SeasonGamemodes.Contains("Hydra Rush"))
                {
                    IXLCell dragonRushCell = winnerInfo.LastDataRowCell.CellRight(2);
                    if (!dragonRushCell.CellLeft().GetString().Equals("Winner"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Winner is not the last line of Dragon Rush");
                    }

                    while (dragonRushCell.GetString().Equals("Nothing"))
                    {
                        if (!dragonRushCell.CellLeft().GetString().Equals("Winner"))
                        {
                            winnerInfo.IsDragonRushRunnerUp = true;
                        }
                        dragonRushCell = dragonRushCell.CellAbove();
                    }
                }

                //If FFA no need to look for teams, else looks for the team
                if (seasonInfo.IsFFA == true)
                {
                    SetSoloWinners(globalStatsList, winList, seasonInfo, winner, seasonLists);
                }
                else
                {
                    SetTeamWinners(globalStatsList, winList, seasonInfo, winnerInfo, winner, seasonLists);
                }

                //Detects double kill runner ups
                winnerInfo.FirstAliveRowCell = winnerInfo.LastDataRowCell.CellRight(2);
                while (winnerInfo.FirstAliveRowCell.CellAbove().GetString().Equals("Nothing"))
                {
                    winnerInfo.FirstAliveRowCell = winnerInfo.FirstAliveRowCell.CellAbove();
                }

                IXLCell lastKillCell = winnerInfo.FirstAliveRowCell.CellAbove();
                IXLCell secondLastKillCell = winnerInfo.FirstAliveRowCell.CellAbove(2);
                if (lastKillCell.GetString().Equals(secondLastKillCell.CellLeft(2).GetString())
                    && secondLastKillCell.GetString().Equals(lastKillCell.CellLeft(2).GetString()))
                {
                    if (seasonInfo.IsFFA == true)
                    {
                        winnerInfo.IsDoubleKillRunnerUp = true;
                    }
                    else
                    {
                        if (!winnerInfo.WinningTeam.Contains(lastKillCell.GetString()) && !winnerInfo.WinningTeam.Contains(secondLastKillCell.GetString()))
                        {
                            winnerInfo.IsDoubleKillRunnerUp = true;
                        }
                    }
                }
            }
            else
            {
                //Check for a double kill ending
                if (winnerInfo.LastDataRowCell.CellRight(2).GetString().Equals(winnerInfo.LastDataRowCell.CellAbove().GetString())
                    && winnerInfo.LastDataRowCell.CellAbove().CellRight(2).GetString().Equals(winnerInfo.LastDataRowCell.GetString()))
                {
                    //Double kill ending so 2 winners
                    winnerInfo.IsDoubleKillWinner = true;
                    String teamWinner = winnerInfo.LastDataRowCell.GetString();
                    String secondTeamWinner = winnerInfo.LastDataRowCell.CellAbove().GetString();
                    winnerInfo.WinnerCell = winnerInfo.LastDataRowCell;
                    winnerInfo.SecondWinnerCell = winnerInfo.LastDataRowCell.CellAbove();

                    //If FFA no need to look for teams, else looks for the team
                    if (seasonInfo.IsFFA == true)
                    {
                        SetSoloWinners(globalStatsList, winList, seasonInfo, teamWinner, seasonLists);
                        SetSoloWinners(globalStatsList, winList, seasonInfo, secondTeamWinner, seasonLists);
                    }
                    else
                    {
                        SetTeamWinners(globalStatsList, winList, seasonInfo, winnerInfo, teamWinner, seasonLists);
                        SetTeamWinners(globalStatsList, winList, seasonInfo, winnerInfo, secondTeamWinner, seasonLists);
                    }
                }
                else
                {
                    //No one won so the Ender Dragon killed everyone
                    winnerInfo.IsDragonWin = true;
                    seasonLists.SeasonWinnerDead.Add("Ender Dragon");
                }
            }
        }
    }
}