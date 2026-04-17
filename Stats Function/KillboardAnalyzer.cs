namespace StatsAnalyzer
{
    public class KillboardAnalyzer
    {
        public static void GetMostKills(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<StatsList> topFragList, 
            List<KillRecords> killRecordsList, SeasonInfo seasonInfo, List<String> seasonTopFrag, List<String> seasonTeams, 
            Dictionary<String, int> killboard, Dictionary<String, int> teamKillboard, String seasonNumberPost)
        {
            if (killboard.Count > 0)
            {
                int topFragAmount = killboard.Values.Max();
                foreach (String killer in killboard.Keys)
                {
                    if (killboard[killer] == topFragAmount)
                    {
                        seasonTopFrag.Add(killer);

                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            GlobalStats.UpdateTopFrags(globalStatsList, killer);
                            topFragList.Add(new StatsList(seasonInfo, killer));
                        }
                    }

                    //Checks if the player that got kills beat their kill record
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        if (killRecordsList.Any(p => p.Player == killer))
                        {
                            KillRecords.UpdateKillRecord(killRecordsList, killer, killboard[killer], seasonInfo);
                        }
                        else
                        {
                            killRecordsList.Add(new KillRecords(killer, killboard[killer], seasonInfo));
                        }
                    }
                }
                RedditPostFormat.FormatTopFrags(redditPosts, seasonTopFrag, topFragAmount, seasonNumberPost);
            }

            //Uses the killboard to make the team killboard
            if (seasonInfo.IsFFA == false)
            {
                foreach (String team in seasonTeams)
                {
                    String[] teamMember = team.Split(',');

                    foreach (String player in teamMember)
                    {
                        if (teamKillboard.ContainsKey(team))
                        {
                            if (killboard.TryGetValue(player, out int kills))
                            {
                                teamKillboard[team] += kills;
                            }
                            else
                            {
                                teamKillboard[team] += 0;
                            }
                        }
                        else
                        {
                            if (killboard.TryGetValue(player, out int kills))
                            {
                                teamKillboard.Add(team, kills);
                            }
                            else
                            {
                                teamKillboard.Add(team, 0);
                            }
                        }
                    }
                }
            }
        }
    }
}