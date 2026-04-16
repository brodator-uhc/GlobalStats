using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class RunnerUpFinder
    {
        public static void SetSoloRunnerUps(List<GlobalStats> globalStatsList, List<StatsList> runnerUpList, SeasonInfo seasonInfo, String runnerUp, 
            List<String> seasonRunnerUps)
        {
            seasonRunnerUps.Add(runnerUp);

            if (seasonInfo.IsCrossoverSeason == false)
            {
                GlobalStats.UpdateRunnerUps(globalStatsList, runnerUp);
                runnerUpList.Add(new StatsList(seasonInfo, runnerUp));
            }
        }
        public static void SetTeamRunnerUps(List<GlobalStats> globalStatsList, List<StatsList> runnerUpList, SeasonInfo seasonInfo, String teamRunnerUp, 
            List<String> seasonRunnerUps, List<String> seasonTeams)
        {
            foreach (String team in seasonTeams)
            {
                if (team.Contains(teamRunnerUp))
                {
                    //Splits the team string to get each player and gives them a runner up
                    String[] runnerUpSplit = team.Split(',');
                    foreach (String runnerUp in runnerUpSplit)
                    {
                        seasonRunnerUps.Add(runnerUp);

                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            GlobalStats.UpdateRunnerUps(globalStatsList, runnerUp);
                            runnerUpList.Add(new StatsList(seasonInfo, runnerUp));
                        }
                    }
                }
            }
        }
        public static void GetRunnerUps(List<GlobalStats> globalStatsList, List<StatsList> runnerUpList, SeasonInfo seasonInfo, WinnerInfo winnerInfo,
            List<String> seasonRunnerUps, List<String> seasonTeams)
        {
            if (winnerInfo.IsDragonWin == false)
            {
                if (winnerInfo.IsDragonRushRunnerUp == false)
                {
                    if (winnerInfo.IsDoubleKillRunnerUp == false)
                    {
                        if (seasonInfo.IsFFA == true)
                        {
                            if (winnerInfo.IsDoubleKillWinner == true)
                            {
                                String runnerUp = winnerInfo.SecondWinnerCell.CellAbove().GetString();
                                SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonRunnerUps);
                            }
                            else
                            {
                                String runnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                                SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonRunnerUps);
                            }
                        }
                        else
                        {
                            if (winnerInfo.IsDoubleKillWinner == true)
                            {
                                //Search for the first cell that does not contain either winning team
                                while (winnerInfo.WinningTeam.Contains(winnerInfo.SecondWinnerCell.CellAbove().GetString()) ||
                                        winnerInfo.SecondWinningTeam.Contains(winnerInfo.SecondWinnerCell.CellAbove().GetString()))
                                {
                                    winnerInfo.SecondWinnerCell = winnerInfo.SecondWinnerCell.CellAbove();
                                }

                                //Figures out the full team of runner ups
                                String teamRunnerUp = winnerInfo.SecondWinnerCell.CellAbove().GetString();
                                SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonRunnerUps, seasonTeams);
                            }
                            else
                            {
                                //Search for the first cell that does not contain the winning team
                                while (winnerInfo.WinningTeam.Contains(winnerInfo.WinnerCell.CellAbove().GetString()))
                                {
                                    winnerInfo.WinnerCell = winnerInfo.WinnerCell.CellAbove();
                                }

                                //Figures out the full team of runner ups
                                String teamRunnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                                SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonRunnerUps, seasonTeams);
                            }
                        }
                    }
                    else
                    {
                        if (seasonInfo.IsFFA == true)
                        {
                            String runnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonRunnerUps);

                            String secondRunnerUp = winnerInfo.WinnerCell.CellAbove(2).GetString();
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, secondRunnerUp, seasonRunnerUps);
                        }
                        else
                        {
                            //Search for the first cell that does not contain the winning team
                            while (winnerInfo.WinningTeam.Contains(winnerInfo.WinnerCell.CellAbove().GetString()))
                            {
                                winnerInfo.WinnerCell = winnerInfo.WinnerCell.CellAbove();
                            }

                            //Figures out the full team of runner ups
                            String teamRunnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                            SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonRunnerUps, seasonTeams);

                            String secondTeamRunnerUp = winnerInfo.WinnerCell.CellAbove(2).GetString();
                            SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, secondTeamRunnerUp, seasonRunnerUps, seasonTeams);
                        }
                    }
                }
                else
                {
                    IXLCell runnerUpPlayer = winnerInfo.LastDataRowCell;
                    IXLCell runnerUpCheck = winnerInfo.LastDataRowCell.CellRight(2);

                    //Goes through everyone that didn't die and adds non winners to the runner ups
                    while (runnerUpCheck.GetString().Equals("Nothing"))
                    {
                        if (!runnerUpCheck.CellLeft().GetString().Equals("Winner"))
                        {
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUpPlayer.GetString(), seasonRunnerUps);
                        }
                        runnerUpCheck = runnerUpCheck.CellAbove();
                        runnerUpPlayer = runnerUpPlayer.CellAbove();
                    }
                }
            }
            else
            {
                //Dragon wins the season, runner up is last dead
                String seasonRunnerUp = winnerInfo.LastDataRowCell.GetString();

                if (seasonInfo.IsFFA == true)
                {
                    SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, seasonRunnerUp, seasonRunnerUps);
                }
                else
                {
                    SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, seasonRunnerUp, seasonRunnerUps, seasonTeams);
                }
            }
        }
    }
}