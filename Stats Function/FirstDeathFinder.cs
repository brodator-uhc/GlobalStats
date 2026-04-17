using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class FirstDeathFinder
    {
        public static void GetFirstDeath(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<StatsList> firstDeathList, 
            SeasonInfo seasonInfo, String seasonNumberPost, IXLCell firstDeathKiller, IXLCell firstDeathVictim)
        {
            String killer = firstDeathKiller.GetString();
            String victim = firstDeathVictim.GetString();
            if (killer.Equals(firstDeathVictim.CellBelow().GetString())
                && victim.Equals(firstDeathKiller.CellBelow().GetString()))
            {
                redditPosts.FirstDeath.Add("**S" + seasonNumberPost + ":** " + victim + " & " + killer + " (Double Kill)" + Environment.NewLine);

                if (seasonInfo.IsCrossoverSeason == false)
                {
                    GlobalStats.UpdateFirstDeaths(globalStatsList, victim);
                    firstDeathList.Add(new StatsList(seasonInfo, victim));

                    GlobalStats.UpdateFirstDeaths(globalStatsList, killer);
                    firstDeathList.Add(new StatsList(seasonInfo, killer));
                }
            }
            else
            {
                if (seasonInfo.IsCrossoverSeason == false)
                {
                    GlobalStats.UpdateFirstDeaths(globalStatsList, victim);
                    firstDeathList.Add(new StatsList(seasonInfo, victim));
                }

                //Double the stats for round exception
                if (seasonInfo.SeasonName.Equals("Game Changer")
                    && seasonInfo.SeasonNumber.Equals("5"))
                {
                    String secondHalf = firstDeathVictim.CellBelow().GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, secondHalf);
                    redditPosts.FirstDeath.Add("**S" + seasonNumberPost + ":** " + victim + " & " + secondHalf + " (" + killer + ")" + Environment.NewLine);
                    firstDeathList.Add(new StatsList(seasonInfo, secondHalf));
                }
                else
                {
                    if (killer.Equals(""))
                    {
                        String method = firstDeathVictim.CellRight(1).GetString();
                        String pveDeath = PveCausesList.GetPveCause(method);
                        redditPosts.FirstDeath.Add("**S" + seasonNumberPost + ":** " + victim + " (" + pveDeath + ")" + Environment.NewLine);
                    }
                    else
                    {
                        redditPosts.FirstDeath.Add("**S" + seasonNumberPost + ":** " + victim + " (" + killer + ")" + Environment.NewLine);
                    }
                }
            }
        }
    }
}