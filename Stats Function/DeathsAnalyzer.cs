using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class DeathsAnalyzer
    {
        public static void GetDeaths(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<RostersList> rostersList, List<RostersList> rostersListNR, 
            List<RoundDebut> roundDebutsList, List<RoundDebut> roundDebutsListNR, List<StatsList> aliveList, List<String> roundRoster, 
            SeasonLists seasonLists, SeasonInfo seasonInfo, IXLRange victimRange, String seasonNumberPost, int firstDataColumn, IXLCell seasonNumberCell)
        {
            if (seasonInfo.IsCrossoverSeason == false)
            {
                //Get seasons data for the all rosters list
                rostersList.Add(new RostersList(seasonInfo));

                if (seasonInfo.IsNR == false)
                {
                    rostersListNR.Add(new RostersList(seasonInfo));
                }
            }

            seasonInfo.SeasonSize = victimRange.RowsUsed().Count();
            foreach (IXLCell deathCell in victimRange.CellsUsed())
            {
                string death = deathCell.GetString();

                if (seasonInfo.IsCrossoverSeason == false)
                {
                    //Checks if its the players debut round, sets the date if it is
                    //If new players sets all the variables for them
                    if (roundDebutsList.Any(r => r.Player == death))
                    {
                        var roundDebutList = roundDebutsList.Find(r => r.Player == death);
                        if (roundDebutList != null)
                        {
                            if (seasonInfo.SeasonDate < roundDebutList.Date)
                            {
                                RoundDebut.UpdateRoundDebut(roundDebutList, seasonInfo);
                            }
                        }
                    }
                    else
                    {
                        roundDebutsList.Add(new RoundDebut(seasonInfo, death));
                        globalStatsList.Add(new GlobalStats(death));
                    }

                    //Checks for players debut round but also excludes non-reddit rounds
                    if (roundDebutsListNR.Any(r => r.Player == death))
                    {
                        if (seasonInfo.IsNR == false)
                        {
                            var roundDebutListNR = roundDebutsListNR.Find(r => r.Player == death);
                            if (roundDebutListNR != null)
                            {
                                if (seasonInfo.SeasonDate < roundDebutListNR.Date)
                                {
                                    RoundDebut.UpdateRoundDebut(roundDebutListNR, seasonInfo);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (seasonInfo.IsNR == false)
                        {
                            roundDebutsListNR.Add(new RoundDebut(seasonInfo, death));
                        }
                    }
                }

                //Add Error Messages for suicides.
                if (deathCell.CellRight(2).GetString().Equals(death))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + death + " suicided! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                }

                //If the players didn't die gets +1 alive on the global stats, else +1 death
                if (deathCell.CellRight(2).GetString().Equals("Nothing"))
                {
                    seasonLists.SeasonAlive.Add(death);
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateAlives(globalStatsList, death);
                        aliveList.Add(new StatsList(seasonInfo, death));
                    }
                }
                else
                {
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateDeaths(globalStatsList, death);
                    }
                }

                //Makes roster for the season, skips players who show up twice with respawns gamemodes
                //Also adds +1 seasons played for the global stats
                if (!seasonLists.SeasonRoster.Contains(death))
                {
                    seasonLists.SeasonRoster.Add(death);
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateSeasonsPlayed(globalStatsList, death);
                    }
                }

                //Makes roster for the round, adds new players
                //Also adds +1 unique round for the global stats
                if (!roundRoster.Contains(death))
                {
                    roundRoster.Add(death);
                    seasonLists.SeasonDebutant.Add(death);
                    GlobalStats.UpdateTotalUniques(globalStatsList, death);
                }
            }

            //Formats the debutants for reddit posts
            seasonLists.SeasonAlive.Sort();
            seasonLists.SeasonDebutant.Sort();
            seasonLists.SeasonRoster.Sort();
            RedditPostFormat.FormatParticipations(redditPosts, seasonLists, seasonNumberPost, firstDataColumn, seasonNumberCell);

            if (seasonInfo.IsCrossoverSeason == false)
            {
                var roundRosterList = rostersList.Find(r => r.Round == seasonInfo.SeasonName && r.Season == seasonInfo.SeasonNumber);
                roundRosterList?.Roster = seasonLists.SeasonRoster;
                //Adds rosters to a list for the sheet, skips non-reddit for the alternate page
                var roundRosterListNR = rostersListNR.Find(r => r.Round == seasonInfo.SeasonName && r.Season == seasonInfo.SeasonNumber);
                roundRosterListNR?.Roster = seasonLists.SeasonRoster;
            }
        }
    }
}