using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class WinnerInfo(IXLWorksheet roundPage)
    {
        public IXLCell WinnerCell { get; set; } = roundPage.Cell(1, 1);
        public IXLCell SecondWinnerCell { get; set; } = roundPage.Cell(1, 1);
        public IXLCell FirstAliveRowCell { get; set; } = roundPage.Cell(1, 1);
        public IXLCell LastDataRowCell { get; set; } = roundPage.Cell(1, 1);
        public String WinningTeam { get; set; } = "";
        public String SecondWinningTeam { get; set; } = "";
        public bool IsDoubleKillWinner { get; set; }
        public bool IsDoubleKillRunnerUp { get; set; }
        public bool IsDragonWin { get; set; }
        public bool IsDragonRushRunnerUp { get; set; }
    }
}