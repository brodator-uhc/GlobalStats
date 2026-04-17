using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class GlobalGapsCalculator
    {
        public static void CalculateGlobalGaps(List<GlobalGaps> globalGapList, IXLWorksheet playerList, IXLWorksheet rosterList)
        {
            //Gets a list of all players.
            Dictionary<String, int> allPlayers = [];
            IXLRange playerRange = playerList.Range(1, 1, playerList.RangeUsed()!.RowCount(), 1);
            foreach (IXLCell player in playerRange.Cells())
            {
                allPlayers.Add(player.Value.ToString(), player.CellRight().GetValue<int>());
            }

            //Calculate if there's a gap between 2 seasons played.
            foreach (String player in allPlayers.Keys)
            {
                String lastRound = "Fake";
                String lastSeason = "1z";
                DateTime lastPlayed = new(2012, 1, 1);
                //Skips players with 1 round played.
                if (allPlayers[player] > 1)
                {
                    //Goes through every seasons and compares gaps of players.
                    IXLRange seasonRange = rosterList.Range(1, 1, rosterList.RangeUsed()!.RowCount(), 1);
                    foreach (IXLCell seasonCell in seasonRange.CellsUsed())
                    {
                        int seasonRow = seasonCell.WorksheetRow().RowNumber();
                        IXLRange rosterRange = rosterList.Range(seasonRow, 4, seasonRow, 129);
                        foreach (IXLCell rosterCell in rosterRange.CellsUsed())
                        {
                            String seasonPlayer = rosterCell.GetString();
                            String seasonRound = rosterList.Cell(seasonRow, 1).Value.ToString();
                            String seasonNumber = rosterList.Cell(seasonRow, 2).Value.ToString();
                            DateTime seasonDate = rosterList.Cell(seasonRow, 3).GetDateTime();

                            if (seasonPlayer == player)
                            {
                                if (lastPlayed != new DateTime(2012, 1, 1))
                                {
                                    TimeSpan timeDiff = seasonDate - lastPlayed;
                                    int gapDays = (int)timeDiff.TotalDays;
                                    if (gapDays > 1095)
                                    {
                                        String playerCompare = player;
                                        foreach (var playerGap in globalGapList.ToList())
                                        {
                                            if (playerGap.DayGap == gapDays &&
                                                playerGap.StartRound == lastRound &&
                                                playerGap.StartSeason == lastSeason &&
                                                playerGap.EndRound == seasonRound &&
                                                playerGap.EndSeason == seasonNumber)
                                            {
                                                playerCompare = playerGap.Player + ", " + player;
                                                globalGapList.Remove(playerGap);
                                            }
                                        }
                                        globalGapList.Add(new GlobalGaps(playerCompare, gapDays, lastRound, lastSeason, lastPlayed, seasonRound, seasonNumber, seasonDate));
                                    }
                                }
                                lastRound = seasonRound;
                                lastSeason = seasonNumber;
                                lastPlayed = seasonDate;
                            }
                        }
                    }
                    Console.WriteLine("Checked gaps for " + player + "!");
                }
            }
        }
    }
}