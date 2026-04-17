using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class RoundGapsCalculator
    {
        public static void CalculateRoundGaps(List<RoundsGaps> roundsGapList, IXLWorksheet roundList, IXLWorksheet rosterList)
        {
            Dictionary<String, int> allRounds = [];
            IXLRange roundRange = roundList.Range(1, 1, roundList.RangeUsed()!.RowCount(), 1);
            foreach (IXLCell round in roundRange.Cells())
            {
                allRounds.Add(round.Value.ToString(), round.CellRight().GetValue<int>());
            }

            foreach (String round in allRounds.Keys)
            {
                String lastRoundSeason = "none";
                Dictionary<String, String> lastSeasonPlayed = [];
                Dictionary<String, DateTime> lastDatePlayed = [];
                //Skips Rounds with 2 or less seasons.
                if (allRounds[round] > 2)
                {
                    //Goes through every season of a round and checks gaps for missed seasons.
                    IXLRange seasonRange = rosterList.Range(1, 1, rosterList.RangeUsed()!.RowCount(), 1);
                    foreach (IXLCell seasonCell in seasonRange.CellsUsed())
                    {
                        int seasonRow = seasonCell.WorksheetRow().RowNumber();
                        String seasonRound = rosterList.Cell(seasonRow, 1).Value.ToString();
                        String seasonNumber = rosterList.Cell(seasonRow, 2).Value.ToString();
                        DateTime seasonDate = rosterList.Cell(seasonRow, 3).GetDateTime();
                        Dictionary<int, String> _roundPlayer = [];
                        Dictionary<int, String> _roundStartSeason = [];
                        Dictionary<int, String> _roundEndSeason = [];
                        Dictionary<int, DateTime> _roundGapDate = [];

                        //Exceptions for crossover rounds with different names.
                        String roundCompare = round;
                        if (seasonRound.Contains(roundCompare))
                        {
                            if (seasonRound == "WMC x Phobia" ||
                                seasonRound == "The Melon Blooded x Scattershot" ||
                                seasonRound == "Phobia x Cinema")
                            {
                                roundCompare = seasonRound;
                            }
                        }

                        if (seasonRound == roundCompare)
                        {
                            IXLRange rosterRange = rosterList.Range(seasonRow, 4, seasonRow, 129);
                            foreach (IXLCell rosterCell in rosterRange.CellsUsed())
                            {
                                String seasonPlayer = rosterCell.GetString();
                                if (!lastRoundSeason.Equals("none"))
                                {
                                    if (lastSeasonPlayed.ContainsKey(seasonPlayer))
                                    {
                                        if (!lastSeasonPlayed[seasonPlayer].Equals(lastRoundSeason))
                                        {
                                            TimeSpan timeDiff = seasonDate - lastDatePlayed[seasonPlayer];
                                            int gapDays = (int)timeDiff.TotalDays;
                                            if (gapDays > 1095)
                                            {
                                                if (_roundPlayer.ContainsKey(gapDays))
                                                {
                                                    _roundPlayer[gapDays] = _roundPlayer[gapDays] + ", " + seasonPlayer;
                                                }
                                                else
                                                {
                                                    _roundPlayer.Add(gapDays, seasonPlayer);
                                                    _roundStartSeason.Add(gapDays, lastSeasonPlayed[seasonPlayer]);
                                                    _roundEndSeason.Add(gapDays, seasonNumber);
                                                    _roundGapDate.Add(gapDays, seasonDate);
                                                }
                                            }
                                        }

                                        lastSeasonPlayed[seasonPlayer] = seasonNumber;
                                        lastDatePlayed[seasonPlayer] = seasonDate;
                                    }
                                    else
                                    {
                                        lastSeasonPlayed.Add(seasonPlayer, seasonNumber);
                                        lastDatePlayed.Add(seasonPlayer, seasonDate);
                                    }
                                }
                                else
                                {
                                    lastSeasonPlayed.Add(seasonPlayer, seasonNumber);
                                    lastDatePlayed.Add(seasonPlayer, seasonDate);
                                }
                            }
                            lastRoundSeason = seasonNumber;

                            foreach (int gap in _roundPlayer.Keys)
                            {
                                roundsGapList.Add(new RoundsGaps(_roundPlayer[gap], gap, roundCompare, _roundStartSeason[gap], _roundEndSeason[gap], _roundGapDate[gap]));
                            }
                        }
                    }
                    Console.WriteLine("Checked gaps for " + round + "!");
                }
            }
        }
    }
}