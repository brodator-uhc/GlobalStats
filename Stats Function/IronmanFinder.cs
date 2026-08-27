using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class IronmanFinder
    {
        public static void GetIronman(List<GlobalStats> globalStatsList, RedditPosts redditPosts, List<StatsList> ironmanList,
            SeasonInfo seasonInfo, String seasonNumberPost, IXLRange ironmanRange, IXLCell ironmanTimeCell)
        {
            String ironmanPost = "**S" + seasonNumberPost + ":** ";
            String ironmanTime = "";
            foreach (IXLCell cell in ironmanRange.CellsUsed())
            {
                string value = cell.GetString();
                if (seasonInfo.IsCrossoverSeason == false)
                {
                    GlobalStats.UpdateIronmans(globalStatsList, value);
                    ironmanList.Add(new StatsList(seasonInfo, value));
                }
                ironmanPost = ironmanPost + value + ", ";
            }
            ironmanPost = ironmanPost[..^2];
            if (ironmanTimeCell.GetString().Equals(""))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("ERROR: Ironman time missing! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
            }
            else
            {
                ironmanTime = " (" + ironmanTimeCell.GetString() + ":" + ironmanTimeCell.CellRight().GetString() + ":" + ironmanTimeCell.CellRight(2).GetString() + ")";
            }
            redditPosts.Ironman.Add(ironmanPost + ironmanTime + Environment.NewLine);
        }
    }
}