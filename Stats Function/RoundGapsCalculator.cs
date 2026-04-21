using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class RoundGapsCalculator
    {
        public static void CalculateRoundGaps(List<RoundsGaps> roundsGapList, IXLWorksheet roundList, IXLWorksheet rosterList)
        {
            Dictionary<String, int> allRounds = [];
            List<GapRosterList> gapRosterLists = [];
            IXLRange roundRange = roundList.Range(1, 1, roundList.RangeUsed()!.RowCount(), 1);
            foreach (IXLCell round in roundRange.Cells())
            {
                allRounds.Add(round.Value.ToString(), round.CellRight().GetValue<int>());
            }

            IXLRange seasonRange = rosterList.Range(1, 1, rosterList.RangeUsed()!.RowCount(), 1);
            foreach (String round in allRounds.Keys)
            {
                gapRosterLists.Add(new GapRosterList(round));
                foreach (IXLCell seasonCell in seasonRange.CellsUsed())
                {
                    int seasonRow = seasonCell.WorksheetRow().RowNumber();
                    String seasonRound = rosterList.Cell(seasonRow, 1).Value.ToString();
                    String seasonNumber = rosterList.Cell(seasonRow, 2).Value.ToString();

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
                            var roundGap = gapRosterLists.Find(r => r.Round == round);
                            roundGap?.Roster.TryAdd(seasonPlayer, seasonNumber);
                        }
                    }
                }
            }

            foreach (String round in allRounds.Keys)
            {
                List<RoundGapSheet> roundGapSheetList = [];
                List<DateTime> seasonDateList = [];
                List<int> seasonGapList = [];
                List<String> seasonNumberList = [];
                var roundGaps = gapRosterLists.Find(r => r.Round == round);
                foreach (String rosterPlayer in roundGaps!.Roster.Keys)
                {
                    roundGapSheetList.Add(new RoundGapSheet(rosterPlayer, roundGaps.Roster[rosterPlayer], roundGapSheetList.Count + 1));
                }
                String lastRoundSeason = "none";
                Dictionary<String, String> lastSeasonPlayed = [];
                Dictionary<String, DateTime> lastDatePlayed = [];
                //Skips Rounds with 2 or less seasons.
                if (allRounds[round] > 2)
                {
                    //Goes through every season of a round and checks gaps for missed seasons.
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
                            if (seasonDateList.Count > 0)
                            {
                                var lastDate = seasonDateList[^1];
                                seasonGapList.Add((int)(seasonDate - lastDate).TotalDays);
                            }
                            seasonDateList.Add(seasonDate);
                            seasonNumberList.Add(seasonNumber);
                            IXLRange rosterRange = rosterList.Range(seasonRow, 4, seasonRow, 129);
                            Dictionary<String, String> _playerSeasonGap = [];
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
                                            _playerSeasonGap.Add(seasonPlayer, gapDays.ToString());
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

                            foreach (String rosterPlayer in roundGaps!.Roster.Keys)
                            {
                                if (lastSeasonPlayed.ContainsKey(rosterPlayer))
                                {
                                    if (lastSeasonPlayed[rosterPlayer] == seasonNumber)
                                    {
                                        var playerRosterGap = roundGapSheetList.Find(p => p.Player == rosterPlayer);
                                        if (_playerSeasonGap.ContainsKey(rosterPlayer))
                                        {
                                            playerRosterGap!.TotalPlayed += 1;
                                            playerRosterGap.Participations.Add(_playerSeasonGap[rosterPlayer]);
                                        }
                                        else
                                        {
                                            playerRosterGap!.TotalPlayed += 1;
                                            playerRosterGap.Participations.Add("x");
                                        }
                                    }
                                    else
                                    {
                                        var playerRosterGap = roundGapSheetList.Find(p => p.Player == rosterPlayer);
                                        playerRosterGap!.Participations.Add("todelete");
                                    }
                                }
                                else
                                {
                                    var playerRosterGap = roundGapSheetList.Find(p => p.Player == rosterPlayer);
                                    playerRosterGap!.Participations.Add("todelete");
                                }
                            }

                            lastRoundSeason = seasonNumber;

                            foreach (int gap in _roundPlayer.Keys)
                            {
                                roundsGapList.Add(new RoundsGaps(_roundPlayer[gap], gap, roundCompare, _roundStartSeason[gap], _roundEndSeason[gap], _roundGapDate[gap]));
                            }
                        }
                    }
                    DataExporter.SaveRoundParticipations(roundGapSheetList, round, seasonDateList, seasonGapList, seasonNumberList);
                    String filePath = "..\\..\\..\\Stats Sheet\\Round Gaps\\" + round + ".xlsx";
                    DataExporter.ClearEmptyCells(filePath);
                    Console.WriteLine("Checked gaps for " + round + "!");
                }
            }
        }
    }
}