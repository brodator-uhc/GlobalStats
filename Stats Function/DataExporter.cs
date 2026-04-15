using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class DataExporter
    {
        public static void ClearEmptyCells()
        {
            var todelete = new XLWorkbook("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
            var playerStatsDoc = todelete.Worksheet(1);

            foreach (var cell in playerStatsDoc.RangeUsed()!.Cells())
            {
                if (cell.Value.ToString() == "todelete")
                {
                    cell.Clear();
                }
            }

            todelete.SaveAs("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
        }

        public static void SaveRoundsGaps(List<RoundsGaps> roundsGap)
        {
            var gapsdoc = new XLWorkbook();
            var gapssheet = gapsdoc.AddWorksheet("Round Gaps");

            gapssheet.Column("F").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapssheet.Cell(1, 1).InsertTable(roundsGap);

            gapssheet.Sort(2, XLSortOrder.Descending);
            gapsdoc.SaveAs("..\\..\\..\\Stats Sheet\\Round Gaps.xlsx");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Round Gaps are now compiled!");
        }
        public static void SaveGlobalGaps(List<GlobalGaps> globalGap)
        {
            var gapsdoc = new XLWorkbook();
            var gapssheet = gapsdoc.AddWorksheet("Global Gaps");

            gapssheet.Column("E").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapssheet.Column("H").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapssheet.Cell(1, 1).InsertTable(globalGap);

            gapssheet.Sort(2, XLSortOrder.Descending);
            gapsdoc.SaveAs("..\\..\\..\\Stats Sheet\\Global Gaps.xlsx");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Player Gaps are now compiled!");
        }

        public static void SavePlayerStats(List<PlayerStats> playerStats)
        {
            var statsdoc = new XLWorkbook();
            var statsheet = statsdoc.AddWorksheet("Player Stats");

            statsheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            statsheet.Cell(1, 1).InsertTable(playerStats);

            statsdoc.SaveAs("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
        }
    }
}