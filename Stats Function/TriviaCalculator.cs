using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class TriviaCalculator
    {
        public static void CalculateTrivia(List<TriviaCount> triviaCounts, IXLWorkbook statsDocument)
        {
            var statsList = statsDocument.Worksheet(1);
            IXLRange methodsRange = statsList.Range(1, 5, statsList.RangeUsed()!.RowCount(), 5);

            foreach (IXLCell methods in methodsRange.CellsUsed())
            {
                if (methods.Value.ToString().Contains("lava") && !methods.Value.ToString().Contains("discovered"))
                {
                    String playerKill = methods.CellRight().Value.ToString();
                    if (playerKill != "Lava")
                    {
                        var lavaKillsStats = triviaCounts.Find(p => p.Player == playerKill);
                        if (lavaKillsStats != null)
                        {
                            lavaKillsStats.Count += 1;
                        }
                        else
                        {
                            triviaCounts.Add(new TriviaCount(playerKill, 1));
                        }
                    }
                }
            }
        }
    }
}