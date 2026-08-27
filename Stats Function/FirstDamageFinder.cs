using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class FirstDamageFinder
    {
        public static void GetFirstDamage(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<StatsList> firstDamageList, 
            SeasonInfo seasonInfo, String seasonNumberPost, IXLRange firstDamageRange, IXLCell firstDamageTimeCell)
        {
            String firstDamagePost = "**S" + seasonNumberPost + ":** ";
            String firstDamageTime = "";
            foreach (IXLCell firstDamageCell in firstDamageRange.CellsUsed())
            {
                string firstDamage = firstDamageCell.GetString();
                if (seasonInfo.IsCrossoverSeason == false)
                {
                    GlobalStats.UpdateFirstDamages(globalStatsList, firstDamage);
                    firstDamageList.Add(new StatsList(seasonInfo, firstDamage));
                }
                firstDamagePost = firstDamagePost + firstDamage + ", ";
            }

            firstDamagePost = firstDamagePost.Remove(firstDamagePost.Length - 2);
            if (firstDamageTimeCell.GetString().Equals(""))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("ERROR: First Damage time missing! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
            }
            else
            {
                firstDamageTime = " (" + firstDamageTimeCell.GetString() + ":" + firstDamageTimeCell.CellRight().GetString() + ":" + firstDamageTimeCell.CellRight(2).GetString() + ")";
            }
            redditPosts.FirstDamage.Add(firstDamagePost + firstDamageTime + Environment.NewLine);
        }
    }
}