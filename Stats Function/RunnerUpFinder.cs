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
            SeasonLists seasonLists)
        {
            foreach (String team in seasonLists.SeasonTeams)
            {
                if (team.Contains(teamRunnerUp))
                {
                    //Splits the team string to get each player and gives them a runner up
                    String[] runnerUpSplit = team.Split(',');
                    foreach (String runnerUp in runnerUpSplit)
                    {
                        seasonLists.SeasonRunnerUp.Add(runnerUp);

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
            SeasonLists seasonLists)
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
                                SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonLists.SeasonRunnerUp);
                            }
                            else
                            {
                                String runnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                                SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonLists.SeasonRunnerUp);
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
                                SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonLists);
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
                                SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonLists);
                            }
                        }
                    }
                    else
                    {
                        if (seasonInfo.IsFFA == true)
                        {
                            String runnerUp = winnerInfo.WinnerCell.CellAbove().GetString();
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUp, seasonLists.SeasonRunnerUp);

                            String secondRunnerUp = winnerInfo.WinnerCell.CellAbove(2).GetString();
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, secondRunnerUp, seasonLists.SeasonRunnerUp);
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
                            SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, teamRunnerUp, seasonLists);

                            String secondTeamRunnerUp = winnerInfo.WinnerCell.CellAbove(2).GetString();
                            SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, secondTeamRunnerUp, seasonLists);
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
                            SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, runnerUpPlayer.GetString(), seasonLists.SeasonRunnerUp);
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
                    SetSoloRunnerUps(globalStatsList, runnerUpList, seasonInfo, seasonRunnerUp, seasonLists.SeasonRunnerUp);
                }
                else
                {
                    SetTeamRunnerUps(globalStatsList, runnerUpList, seasonInfo, seasonRunnerUp, seasonLists);
                }
            }
        }
    }
}