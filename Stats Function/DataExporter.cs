using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class DataExporter
    {
        public static void ClearEmptyCells()
        {
            var toDelete = new XLWorkbook("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
            var playerStatsDoc = toDelete.Worksheet(1);

            foreach (var cell in playerStatsDoc.RangeUsed()!.Cells())
            {
                if (cell.Value.ToString() == "todelete")
                {
                    cell.Clear();
                }
            }

            toDelete.SaveAs("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
        }

        public static void SaveRoundsGaps(List<RoundsGaps> roundsGap)
        {
            var gapsDoc = new XLWorkbook();
            var gapsSheet = gapsDoc.AddWorksheet("Round Gaps");

            gapsSheet.Column("F").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapsSheet.Cell(1, 1).InsertData(roundsGap);

            gapsSheet.Sort(2, XLSortOrder.Descending);
            gapsDoc.SaveAs("..\\..\\..\\Stats Sheet\\Round Gaps.xlsx");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Round Gaps are now compiled!");
        }
        public static void SaveGlobalGaps(List<GlobalGaps> globalGap)
        {
            var gapsDoc = new XLWorkbook();
            var gapsSheet = gapsDoc.AddWorksheet("Global Gaps");

            gapsSheet.Column("E").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapsSheet.Column("H").Style.NumberFormat.Format = "mm/dd/yyyy";
            gapsSheet.Cell(1, 1).InsertData(globalGap);

            gapsSheet.Sort(2, XLSortOrder.Descending);
            gapsDoc.SaveAs("..\\..\\..\\Stats Sheet\\Global Gaps.xlsx");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Player Gaps are now compiled!");
        }

        public static void SavePlayerStats(List<PlayerStats> playerStats)
        {
            var statsDoc = new XLWorkbook();
            var statSheet = statsDoc.AddWorksheet("Player Stats");

            statSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            statSheet.Cell(1, 1).InsertData(playerStats);

            statsDoc.SaveAs("..\\..\\..\\Stats Sheet\\Player Stats.xlsx");
        }

        public static void SaveRoundList(List<RoundList> roundLists, List<PveCausesList> pveCausesLists, List<GamemodesList> gamemodesLists, List<RostersList> rostersLists, List<RostersList> rostersListsNR, String postFolder)
        {
            var roundListDoc = new XLWorkbook();

            //Making Round List Page
            var roundListSheet = roundListDoc.AddWorksheet("Round List");
            roundListSheet.Column("D").Style.NumberFormat.Format = "dd mmm, yyyy";
            roundListSheet.Cell(1, 1).InsertData(roundLists);
            roundListSheet.Sort(4, XLSortOrder.Ascending);

            //Making PvE List Page
            var pveListSheet = roundListDoc.AddWorksheet("PvE List");
            pveListSheet.Cell(1, 1).InsertData(pveCausesLists);
            pveListSheet.Sort(2, XLSortOrder.Descending);

            //Making Gamemode List Page
            var gamemodeListSheet = roundListDoc.AddWorksheet("Gamemode List");
            gamemodeListSheet.Cell(1, 1).InsertData(gamemodesLists);
            gamemodeListSheet.Sort(2, XLSortOrder.Descending);

            //Making All Rosters Page
            var allRostersSheet = roundListDoc.AddWorksheet("All Rosters");
            allRostersSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            int currentRow = 1;
            foreach (var round in rostersLists)
            {
                allRostersSheet.Cell(currentRow, 1).Value = round.Round;
                allRostersSheet.Cell(currentRow, 2).Value = round.Season;
                allRostersSheet.Cell(currentRow, 3).Value = round.Date;
                allRostersSheet.Cell(currentRow, 4).InsertData(round.Roster, transpose: true);
                currentRow++;
            }
            allRostersSheet.Sort(3, XLSortOrder.Ascending);

            //Making NR All Rosters Page
            var allRostersNRSheet = roundListDoc.AddWorksheet("All Rosters (NR)");
            allRostersNRSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            int currentRowNR = 1;
            foreach (var round in rostersListsNR)
            {
                allRostersNRSheet.Cell(currentRowNR, 1).Value = round.Round;
                allRostersNRSheet.Cell(currentRowNR, 2).Value = round.Season;
                allRostersNRSheet.Cell(currentRowNR, 3).Value = round.Date;
                allRostersNRSheet.Cell(currentRowNR, 4).InsertData(round.Roster, transpose: true);
                currentRowNR++;
            }
            allRostersNRSheet.Sort(3, XLSortOrder.Ascending);

            roundListDoc.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Round_List.xlsx");
        }

        public static void SaveRoundDebut(List<RoundDebut> roundDebutsList, List<RoundDebut> roundDebutsListNR, String postFolder)
        {
            var roundDebutDoc = new XLWorkbook();

            //RR Debuts
            var roundDebutSheet = roundDebutDoc.AddWorksheet("RR Debuts");
            roundDebutSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            roundDebutSheet.Cell(1, 1).InsertData(roundDebutsList);
            roundDebutSheet.Sort(3, XLSortOrder.Ascending);

            //RR Debuts (No NR)
            var roundDebutSheetNR = roundDebutDoc.AddWorksheet("RR Debuts (No NR)");
            roundDebutSheetNR.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            roundDebutSheetNR.Cell(1, 1).InsertData(roundDebutsListNR);
            roundDebutSheetNR.Sort(3, XLSortOrder.Ascending);

            roundDebutDoc.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\RR_Debuts.xlsx");
        }

        public static void SaveGlobalStats(List<GlobalStats> globalStatsList, List<KillRecords> killRecordsList, String postFolder)
        {
            var globalStatsDoc = new XLWorkbook();

            //Global Stats
            var globalStatsSheet = globalStatsDoc.AddWorksheet("Global Stats");
            globalStatsSheet.Cell(1, 1).InsertData(globalStatsList);
            globalStatsSheet.Sort(1, XLSortOrder.Ascending);

            //Kill Records
            var killRecordSheet = globalStatsDoc.AddWorksheet("Kill Records");
            killRecordSheet.Cell(1, 1).InsertData(killRecordsList);
            killRecordSheet.Sort(1, XLSortOrder.Ascending);

            globalStatsDoc.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Global_Stats.xlsx");
        }

        public static void SaveCompiledStats(List<KillsList> killsLists, List<TeamsList> teamsLists, List<StatsList> firstDamageLists, List<StatsList> ironmanLists, List<StatsList> pveDeathLists, List<StatsList> firstDeathLists, List<StatsList> firstBloodLists, List<StatsList> topFragLists, List<StatsList> runnerUpLists, List<StatsList> aliveLists, List<StatsList> winLists, String postFolder)
        {
            var compiledStatsDoc = new XLWorkbook();

            //Making Kills Page
            var compiledKillsSheet = compiledStatsDoc.AddWorksheet("All Kills");
            compiledKillsSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledKillsSheet.Cell(1, 1).InsertData(killsLists);
            compiledKillsSheet.Sort(3, XLSortOrder.Ascending);

            //Making Teams Page
            var compiledTeamsSheet = compiledStatsDoc.AddWorksheet("All Teams");
            compiledTeamsSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledTeamsSheet.Cell(1, 1).InsertData(teamsLists);
            compiledTeamsSheet.Sort(3, XLSortOrder.Ascending);

            //First Damage list
            var compiledFirstDamageSheet = compiledStatsDoc.AddWorksheet("First Damage");
            compiledFirstDamageSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledFirstDamageSheet.Cell(1, 1).InsertData(firstDamageLists);
            compiledFirstDamageSheet.Sort(3, XLSortOrder.Ascending);

            //Ironman list
            var compiledIronmanSheet = compiledStatsDoc.AddWorksheet("Ironman");
            compiledIronmanSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledIronmanSheet.Cell(1, 1).InsertData(ironmanLists);
            compiledIronmanSheet.Sort(3, XLSortOrder.Ascending);

            //PvE Death list
            var compiledPveSheet = compiledStatsDoc.AddWorksheet("PvE Deaths");
            compiledPveSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledPveSheet.Cell(1, 1).InsertData(pveDeathLists);
            compiledPveSheet.Sort(3, XLSortOrder.Ascending);

            //First Death list
            var compiledFirstDeathSheet = compiledStatsDoc.AddWorksheet("First Death");
            compiledFirstDeathSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledFirstDeathSheet.Cell(1, 1).InsertData(firstDeathLists);
            compiledFirstDeathSheet.Sort(3, XLSortOrder.Ascending);

            //First Blood list
            var compiledFirstBloodSheet = compiledStatsDoc.AddWorksheet("First Blood");
            compiledFirstBloodSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledFirstBloodSheet.Cell(1, 1).InsertData(firstBloodLists);
            compiledFirstBloodSheet.Sort(3, XLSortOrder.Ascending);

            //Most Kills list
            var compiledTopFragsSheet = compiledStatsDoc.AddWorksheet("Top Frags");
            compiledTopFragsSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledTopFragsSheet.Cell(1, 1).InsertData(topFragLists);
            compiledTopFragsSheet.Sort(3, XLSortOrder.Ascending);

            //Runner Up list
            var compiledRunnerUpSheet = compiledStatsDoc.AddWorksheet("Runner Ups");
            compiledRunnerUpSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledRunnerUpSheet.Cell(1, 1).InsertData(runnerUpLists);
            compiledRunnerUpSheet.Sort(3, XLSortOrder.Ascending);

            //Alive list
            var compiledAliveSheet = compiledStatsDoc.AddWorksheet("Alive");
            compiledAliveSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledAliveSheet.Cell(1, 1).InsertData(aliveLists);
            compiledAliveSheet.Sort(3, XLSortOrder.Ascending);

            //Win list
            var compiledWinSheet = compiledStatsDoc.AddWorksheet("Wins");
            compiledWinSheet.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
            compiledWinSheet.Cell(1, 1).InsertData(winLists);
            compiledWinSheet.Sort(3, XLSortOrder.Ascending);

            compiledStatsDoc.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Stats_Compiled.xlsx");
        }
    }
}