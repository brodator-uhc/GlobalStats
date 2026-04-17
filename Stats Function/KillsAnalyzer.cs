using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class KillsAnalyzer
    {
        public static void GetKills(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<KillsList> killsList, 
            List<StatsList> firstBloodList, List<StatsList> pveDeathList, List<PveCausesList> pveCausesList, SeasonInfo seasonInfo, 
            Dictionary<String, int> killboard, String seasonNumberPost, IXLRange killerRange)
        {
            bool firstBlood = false;
            foreach (IXLCell killCell in killerRange.Cells())
            {
                String killer = killCell.GetString();
                String victim = killCell.CellLeft(2).GetString();
                String method = killCell.CellLeft().GetString();

                //Checks if killer is PvE or Player
                if (globalStatsList.Any(p => p.Player == killer))
                {
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        killsList.Add(new KillsList(seasonInfo, victim, method, killer));
                        GlobalStats.UpdateKills(globalStatsList, killer);
                    }

                    RedditPostsKills.UpdateKills(redditPosts, killer, victim, seasonNumberPost);

                    //Figures out the killboard of the season
                    if (!killboard.TryAdd(killer, 1))
                    {
                        killboard[killer] += 1;
                    }

                    //Check if there was a double kill for first blood, otherwise gives it to the first player found
                    if (firstBlood == false)
                    {
                        if (killer.Equals(killCell.CellBelow().CellLeft(2).GetString())
                            && killCell.CellBelow().GetString().Equals(killCell.CellLeft(2).GetString()))
                        {
                            firstBlood = true;
                            String secondKiller = killCell.CellBelow().GetString();
                            redditPosts.FirstBlood.Add("**S" + seasonNumberPost + ":** " + killer + " & " + secondKiller + " (Double Kill)" + Environment.NewLine);

                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);
                                firstBloodList.Add(new StatsList(seasonInfo, killer));

                                GlobalStats.UpdateFirstBloods(globalStatsList, secondKiller);
                                firstBloodList.Add(new StatsList(seasonInfo, secondKiller));
                            }
                        }
                        else
                        {
                            firstBlood = true;
                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);
                                firstBloodList.Add(new StatsList(seasonInfo, killer));
                            }

                            //Double the stats for round exception
                            if (seasonInfo.SeasonName.Equals("Game Changer")
                                && seasonInfo.SeasonNumber.Equals("5"))
                            {
                                String secondHalf = killCell.CellBelow().GetString();
                                String secondHalfKill = killCell.CellBelow().CellLeft(2).GetString();

                                redditPosts.FirstBlood.Add("**S" + seasonNumberPost + ":** " + killer + " & " + secondHalf + " (" + victim + " & " + secondHalfKill + ")" + Environment.NewLine);
                                
                                GlobalStats.UpdateFirstBloods(globalStatsList, secondHalf);
                                firstBloodList.Add(new StatsList(seasonInfo, secondHalf));
                            }
                            else
                            {
                                redditPosts.FirstBlood.Add("**S" + seasonNumberPost + ":** " + killer + " (" + victim + ")" + Environment.NewLine);
                            }
                        }
                    }
                }
                else
                {
                    //Adds +1 PvE Death for the player
                    if (!killer.Equals("Nothing"))
                    {
                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            String pveVictim = victim;
                            GlobalStats.UpdatePveDeaths(globalStatsList, pveVictim);
                            pveDeathList.Add(new StatsList(seasonInfo, pveVictim));
                        }

                        //Filters all the unique pve deaths
                        if (killer.Equals(""))
                        {
                            String pvedeath = PveCausesList.GetPveCause(method);

                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                killsList.Add(new KillsList(seasonInfo, victim, method, pvedeath));

                                if (pveCausesList.Any(p => p.PveCause == pvedeath))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, pvedeath);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(pvedeath, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, pvedeath, victim, seasonNumberPost);
                        }
                        else
                        {
                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                killsList.Add(new KillsList(seasonInfo, victim, method, killer));

                                if (pveCausesList.Any(p => p.PveCause == killer))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, killer);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(killer, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, killer, victim, seasonNumberPost);
                        }
                    }
                }
            }
        }
    }
}